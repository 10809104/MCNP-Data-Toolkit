/* ================= KikKoh @2026 =================
   ================ 最終版：支援複選資料夾批次處理 ============ */
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <shellapi.h>
#include <dirent.h>
#include <commdlg.h>
#include <time.h>

#include "config.h"
#include "tools.h"

#define MAX_PATH 260

// 定義 UI 元件 ID
#define ID_LISTBOX  101
#define ID_BTN_OK   102
#define ID_BTN_HELP 103

// ===== 顏色定義 (Console Color Codes) =====
#define LOG_INFO    11 // 淺藍
#define LOG_SUCCESS 10 // 綠色
#define LOG_WARN    14 // 黃色
#define LOG_ERROR   12 // 紅色
#define LOG_SYSTEM  9  // 藍色
#define LOG_NORMAL  7  // 預設白色

// --- 全域變數 ---
HWND hListBox, hBtnOk;
HFONT hFont;
char G_SelectedSourcePath[MAX_PATH] = {0}; // 存儲來源 CSV 路徑
HANDLE hConsole;

/**
 * 變更 Console 輸出顏色
 */
void setColor(int color) {
    SetConsoleTextAttribute(hConsole, color);
}

/**
 * 格式化輸出日誌
 * @param color 顏色代碼
 * @param tag 標籤文字 (例如 INFO, WARN)
 * @param fmt 格式化字串
 */
void logMessage(int color, const char* tag, const char* fmt, ...) {
    setColor(color);
    printf("[%s] ", tag);

    va_list args;
    va_start(args, fmt);
    vprintf(fmt, args);
    va_end(args);

    printf("\n");
    setColor(LOG_NORMAL);
}

// --- 視窗輔助函式 ---
/**
 * 設定視窗元件為微軟正黑體
 */
void SetDefaultFont(HWND hwnd) {
    hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                       VARIABLE_PITCH | FF_SWISS, "Microsoft JhengHei");
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

/**
 * 開啟檔案選取對話框
 */
int SelectSourceFile(HWND hwnd, char* outPath) {
    OPENFILENAME ofn;
    char szFile[MAX_PATH] = {0};
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "only CSV Files (xxx.csv)\0*.csv\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;
	if (GetOpenFileName(&ofn)) {
	    // 1. 先取得副檔名位置
	    char *ext = strrchr(ofn.lpstrFile, '.');
	    
	    // 2. 檢查副檔名是否存在，且是否為 .csv (不分大小寫)
	    if (ext != NULL && _stricmp(ext, ".csv") == 0) {
	        strcpy(outPath, ofn.lpstrFile);
	        return 1;
	    } else {
	        // 如果選錯了，彈出警告
	        MessageBox(hwnd, "錯誤：請選擇有效的 CSV 檔案！", "格式不符", MB_OK | MB_ICONERROR);
	        return 0;
	    }
	}
    return 0;
}

// ===== 核心批次處理邏輯 =====
void ExecuteBatchProcess(HWND hwnd) {
    ULONGLONG total_start = GetTickCount(); // 計時開始
    int total_processed_files = 0;
    int total_directories = 0;

    // 取得 ListBox 選取的項目數量
    int count = SendMessage(hListBox, LB_GETSELCOUNT, 0, 0);
    if (count <= 0) {
        MessageBox(hwnd, "請至少選擇一個資料夾！", "提示", MB_OK | MB_ICONWARNING);
        return;
    }

    // 配置記憶體存儲索引值
    int *indices = (int *)malloc(sizeof(int) * count);
    if (!indices) {
        logMessage(LOG_ERROR, "ERROR", "Memory allocation failed for indices");
        return;
    }

    SendMessage(hListBox, LB_GETSELITEMS, (WPARAM)count, (LPARAM)indices);
    logMessage(LOG_SYSTEM, "SYSTEM", "Batch process started (%d directories)", count);

    for (int i = 0; i < count; i++) {
        ULONGLONG folder_start = GetTickCount();
        total_directories++;

        char folder[MAX_PATH];
        SendMessage(hListBox, LB_GETTEXT, indices[i], (LPARAM)folder);
        logMessage(LOG_INFO, "INFO", "[%d/%d] Processing directory: %s", i + 1, count, folder);

        DynamicTable origin;
        FileInfo file_info = {0, NULL, 0, 0, 0};
        TableSet* myData = NULL;
        char **myFiles = NULL;

        // 載入原始對照表
        int rows = loadSourceCSV(G_SelectedSourcePath, &origin);
        if (rows <= 0) {
            logMessage(LOG_WARN, "WARNING", "Failed to read source CSV: %s. Skipping folder.", G_SelectedSourcePath);
            continue;
        }

        // 掃描目錄下的 .o 檔案
        DIR *dir = opendir(folder);
        if (!dir) {
            logMessage(LOG_WARN, "WARNING", "Directory not found: %s. Skipping.", folder);
            clearTable(&origin);
            continue;
        }

        int total_o_files = 0;
        struct dirent *ent;
        while ((ent = readdir(dir)) != NULL) {
            char *ext = strrchr(ent->d_name, '.');
            if (ext && strcmp(ext, ".o") == 0) {
                total_o_files++;
                if (total_o_files == 1) { // 讀取第一個檔案作為結構基準
                    char full_path[512];
                    snprintf(full_path, sizeof(full_path), "%s/%s", folder, ent->d_name);
                    file_info = get_file_info(full_path);
                }
            }
        }
        closedir(dir);
        total_processed_files += total_o_files;

        if (total_o_files == 0 || file_info.particles == 0 || file_info.total_count == 0) {
            logMessage(LOG_WARN, "WARNING", "No valid .o files or tally data. Skipping folder.");
            clearTable(&origin);
            continue;
        }

        // 排序並載入資料
        total_o_files = get_sorted_o_files(folder, &myFiles);
        myData = createTableSet(total_o_files,
                                file_info.total_count + 1,
                                file_info.label_count + 1);

        if (!myData) {
            logMessage(LOG_ERROR, "ERROR", "Failed to allocate TableSet for folder: %s", folder);
            clearTable(&origin);
            if (myFiles) cleanup_list(myFiles, total_o_files);
            continue;
        }

        // 遍歷處理單一檔案
        for (int j = 0; j < total_o_files; j++) {
            char full_path[512];
            snprintf(full_path, sizeof(full_path), "%s/%s", folder, myFiles[j]);
            logMessage(LOG_INFO, "INFO", "Processing file [%d/%d]: %s", j + 1, total_o_files, myFiles[j]);

            loadSingleFileData(full_path,
                               &(myData->tables[j]),
                               file_info.particles,
                               file_info.tally_count,
                               file_info.label_count,
                               file_info.labels);
        }

        // 導出報表
        exportFinalReport(folder, &origin, myData, myFiles);
        logMessage(LOG_SUCCESS, "SUCCESS", "Folder processed successfully: %s", folder);

        // 清理記憶體
        clearTable(&origin);
        if (myFiles) cleanup_list(myFiles, total_o_files);
        if (myData) freeTableSet(myData);
        if (file_info.labels) {
            for (int k = 0; k < file_info.label_count; k++)
                if (file_info.labels[k]) free(file_info.labels[k]);
            free(file_info.labels);
        }

        ULONGLONG folder_end = GetTickCount();
        double folder_time = (folder_end - folder_start) / 1000.0;
        logMessage(LOG_SYSTEM, "SYSTEM", "Folder elapsed time: %.2f seconds", folder_time);
    }

    free(indices);

    ULONGLONG total_end = GetTickCount();
    double total_time = (total_end - total_start) / 1000.0;

    // 結算畫面
    printf("\n");
    logMessage(LOG_SYSTEM, "SUMMARY", "Total directories processed : %d", total_directories);
    logMessage(LOG_SYSTEM, "SUMMARY", "Total .o files processed     : %d", total_processed_files);
    logMessage(LOG_SYSTEM, "SUMMARY", "Total execution time         : %.2f seconds", total_time);
    printf("\n");

    // 整合功能詢問
    int msgboxID = MessageBox(
        hwnd,
        "批次 CSV 處理完成！\n是否要立即執行 PowerShell 整合為分頁 Excel？",
        "批次處理完成",
        MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON1
    );

    if (msgboxID == IDYES) {
        logMessage(LOG_SYSTEM, "SYSTEM", "Launching PowerShell integration script...");
        ShellExecute(NULL, "open", "powershell.exe",
                     "-ExecutionPolicy Bypass -File merge.ps1",
                     NULL, SW_HIDE);

        MessageBox(hwnd,
                   "Excel 整合指令已送出，請檢查 Total_Summary_Styled.xlsx。",
                   "完成",
                   MB_OK | MB_ICONINFORMATION);
    }
}

// --- 視窗回呼函式 ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            // 建立 UI 佈局
            HWND hGroup = CreateWindow("BUTTON", " 目錄複選 (Ctrl/Shift) ", 
                                       WS_VISIBLE | WS_CHILD | BS_GROUPBOX, 
                                       10, 10, 265, 180, hwnd, NULL, NULL, NULL);
            SetDefaultFont(hGroup);

            hListBox = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", NULL, 
                                    WS_VISIBLE | WS_CHILD | LBS_STANDARD | WS_VSCROLL | LBS_EXTENDEDSEL, 
                                    25, 40, 235, 130, hwnd, (HMENU)ID_LISTBOX, NULL, NULL);
            SetDefaultFont(hListBox);

            HWND hBtnHelp = CreateWindow("BUTTON", "說明 (Help)", 
                                         WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 
                                         15, 205, 80, 35, hwnd, (HMENU)ID_BTN_HELP, NULL, NULL);
            SetDefaultFont(hBtnHelp);

            hBtnOk = CreateWindow("BUTTON", "批次處理", 
                                       WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 
                                       155, 205, 120, 35, hwnd, (HMENU)ID_BTN_OK, NULL, NULL);
            SetDefaultFont(hBtnOk);

            // 自動掃描當前目錄下的資料夾並加入 ListBox
            WIN32_FIND_DATA findData;
            HANDLE hFind = FindFirstFile(".\\*", &findData);
            if (hFind != INVALID_HANDLE_VALUE) {
                do {
                    if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && findData.cFileName[0] != '.') {
                        SendMessage(hListBox, LB_ADDSTRING, 0, (LPARAM)findData.cFileName);
                    }
                } while (FindNextFile(hFind, &findData));
                FindClose(hFind);
            }
            break;
        }
        case WM_COMMAND:
            if (LOWORD(wParam) == ID_BTN_OK) ExecuteBatchProcess(hwnd);
            if (LOWORD(wParam) == ID_BTN_HELP) {
                MessageBox(hwnd, 
                    "使用說明：\n"
                    "1. 目錄結構\n"
                    "/file/\n"
					" ├──OO.exe\n"
					" ├──merge.ps1\n"
					" ├──(source.csv)\n"
					" ├──(files)\n"
					" │        └──OO.o\n"
                    "2. 選取 Source CSV。\n"
                    "3. 在清單中多選資料夾 (Ctrl/Shift)。\n"
                    "4. 點擊「批次處理」產出數據。\n"
					"5. 最後可選擇整合為單一 Excel 檔。", 
                    "說明資訊", MB_OK | MB_ICONINFORMATION);
            }
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// --- 主程式進入點 ---
int main() {
    AllocConsole();
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    freopen("CONOUT$", "w", stdout);
    
    SetConsoleTitle("卷卷又糖糖");
	
    // 神獸守護畫面
    setColor(LOG_SUCCESS);
	printf("                /\\_/\\      ___      /\\_/\\\n");
	printf("               ( o.o )   /   \\    ( o.o )\n");
	printf("                > ^ <   / /^\\ \\    > ^ <\n");
	printf("             ___/     \\_/ /_\\ \\_/     \\___\n");
	printf("            /  /|       (  * )       |\\  \\\n");
	printf("           /  / |        \\___/        | \\  \\\n");
	printf("          (  (  |      Protect Code!  |  )  )\n");
	printf("           \\  \\ |     Keep Bug Free    | /  /\n");
	printf("            \\__\\|_____________________|/__/\n");
	printf("                    神獸守護，程式不閃退!\n");
	
    // 標題裝飾
    setColor(LOG_INFO);
    printf(" __________________________________________________________ \n");
    printf("|                                                          |\n");
    printf("|         M R C P   D A T A   T O O L K I T   v2.7         |\n");
    printf("|__________________________________________________________|\n\n");
    setColor(LOG_NORMAL);
	
	// 檢查更新 
	CheckForUpdates();

    // 檢查預設檔案
    FILE *f = fopen("source.csv", "r");
    if (f) {
        logMessage(LOG_INFO, "INFO", "Default 'source.csv' detected.");
        strncpy(G_SelectedSourcePath, "source.csv", sizeof(G_SelectedSourcePath)-1);
		G_SelectedSourcePath[sizeof(G_SelectedSourcePath)-1] = '\0';
        fclose(f);
    } else {
        logMessage(LOG_WARN, "WARN", "'source.csv' missing. Requesting user input...");
        MessageBox(NULL, "偵測不到 source.csv 檔案！\n請點擊確定後手動選取來源 CSV。", "找不到檔案", MB_OK | MB_ICONINFORMATION);
        
        if (!SelectSourceFile(NULL, G_SelectedSourcePath)) {
            logMessage(LOG_ERROR, "FATAL", "User cancelled selection. Exiting.");
            MessageBox(NULL, "未選取來源檔案，程式將結束。", "Error", MB_OK | MB_ICONERROR);
            return 0;
        }
    }

    logMessage(LOG_SUCCESS, "READY", "Current source: %s", G_SelectedSourcePath);
    printf(" Waiting for user interaction...\n\n");

    // 註冊視窗類別與建立視窗
    HINSTANCE hInstance = GetModuleHandle(NULL);
    WNDCLASS wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = "MRCPApp";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClass(&wc);

    HWND hwnd = CreateWindow("MRCPApp", "MRCP 數據整合工具", 
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_VISIBLE, 
                             (GetSystemMetrics(SM_CXSCREEN)-300)/2, (GetSystemMetrics(SM_CYSCREEN)-300)/2, 
                             300, 300, NULL, NULL, hInstance, NULL);

    // 訊息迴圈
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!IsDialogMessage(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // --- 結束前的倒數邏輯 (美化版) ---
    printf("\n");
    setColor(LOG_SYSTEM);
    printf("============================================================\n");
    setColor(LOG_NORMAL);
    printf(" 感謝您的使用。\n");
    
    for(int i = 3; i > 0; i--) {
        printf(" 程式將在 ");
        setColor(LOG_WARN);
        printf("%d", i);
        setColor(LOG_NORMAL);
        printf(" 秒後自動關閉...\r");
        fflush(stdout);
        Sleep(1000);
    }
    
    setColor(LOG_SYSTEM);
    printf("\n============================================================\n");
    setColor(LOG_NORMAL);

    FreeConsole();
    return 0;
}
