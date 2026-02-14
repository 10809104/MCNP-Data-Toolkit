/* ================= KikKoh @2026 =================
   ================ 最終版：支援複選資料夾批次處理 ============ */
#define UNICODE
#define _UNICODE

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
 * 開啟檔案選取對話框 (Unicode 版)
 */
int SelectSourceFile(HWND hwnd, WCHAR* outPath) {
    OPENFILENAMEW ofn;
    WCHAR szFile[MAX_PATH] = {0};
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile)/sizeof(WCHAR);
    ofn.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetOpenFileNameW(&ofn)) {
        WCHAR *ext = wcsrchr(ofn.lpstrFile, L'.');
        if (ext && _wcsicmp(ext, L".csv") == 0) {
            wcscpy(outPath, ofn.lpstrFile);
            return 1;
        } else {
            MessageBoxW(hwnd, L"錯誤：請選擇有效的 CSV 檔案！", L"格式不符", MB_OK | MB_ICONERROR);
            return 0;
        }
    }
    return 0;
}

// --- 自定義按鈕文字的 Hook 邏輯 ---
HHOOK hMsgBoxHook;

/** 自訂 MessageBox 按鈕文字 Hook (Unicode) */
LRESULT CALLBACK MsgBoxHookProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HCBT_ACTIVATE) {
        HWND hwndMsgBox = (HWND)wParam;
        if (GetDlgItem(hwndMsgBox, IDOK)) SetDlgItemTextW(hwndMsgBox, IDOK, L"萬歲！");
        if (GetDlgItem(hwndMsgBox, IDCANCEL)) SetDlgItemTextW(hwndMsgBox, IDCANCEL, L"我要延畢...");
        return 0;
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

/** 顯示畢業視窗 */
int ShowGraduationDialog(HWND hwnd) {
    HHOOK hMsgBoxHook = SetWindowsHookExW(WH_CBT, MsgBoxHookProc, NULL, GetCurrentThreadId());
    int result = MessageBoxW(hwnd, L"恭喜操作完成，可以順利畢業了！", L"順利執行完畢", MB_OKCANCEL | MB_ICONINFORMATION);
    UnhookWindowsHookEx(hMsgBoxHook);
    return result;
}

// ===== 核心批次處理邏輯 =====
void ExecuteBatchProcess(HWND hwnd) {
    ULONGLONG total_start = GetTickCount();
    int total_processed_files = 0;
    int total_directories = 0;

    // 取得 ListBox 選取的項目數量
    int count = SendMessage(hListBox, LB_GETSELCOUNT, 0, 0);
    if (count <= 0) {
        MessageBox(hwnd, "請至少選擇一個資料夾！", "提示", MB_OK | MB_ICONWARNING);
        return;
    }

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

        // --- 1. 資源初始化為 NULL / 零值 ---
        // 這是 goto 清理模式的核心：確保清理時知道哪些資源已分配
        DynamicTable origin = {0}; 
        FileInfo file_info = {0, NULL, 0, 0, 0};
        TableSet* myData = NULL;
        char **myFiles = NULL;
        int total_o_files = 0;
        DIR *dir = NULL;

        char folder[MAX_PATH];
        SendMessage(hListBox, LB_GETTEXT, indices[i], (LPARAM)folder);
        logMessage(LOG_INFO, "INFO", "[%d/%d] Processing directory: %s", i + 1, count, folder);

        // --- 2. 載入原始對照表 ---
        int rows = loadSourceCSV(G_SelectedSourcePath, &origin);
        if (rows <= 0) {
            logMessage(LOG_WARN, "WARNING", "Failed to read source CSV: %s", G_SelectedSourcePath);
            goto folder_cleanup; // 直接跳到清理區
        }

        // --- 3. 掃描目錄 ---
        dir = opendir(folder);
        if (!dir) {
            logMessage(LOG_WARN, "WARNING", "Directory not found: %s", folder);
            goto folder_cleanup;
        }

        struct dirent *ent;
        while ((ent = readdir(dir)) != NULL) {
            char *ext = strrchr(ent->d_name, '.');
            if (ext && _stricmp(ext, ".o") == 0) {
                total_o_files++;
                if (total_o_files == 1) {
                    char full_path[512];
                    snprintf(full_path, sizeof(full_path), "%s\\%s", folder, ent->d_name);
                    file_info = get_file_info(full_path);
                }
            }
        }
        closedir(dir);
        dir = NULL; // 標記已關閉

        if (total_o_files <= 0 || file_info.total_count <= 0) {
            logMessage(LOG_WARN, "WARNING", "No valid .o files in: %s", folder);
            goto folder_cleanup;
        }

        // --- 4. 排序並配置記憶體 ---
        total_o_files = get_sorted_o_files(folder, &myFiles);
        if (total_o_files <= 0) goto folder_cleanup;

        myData = createTableSet(total_o_files, 
                               file_info.total_count + 1, 
                               file_info.label_count + 1);
        if (!myData) {
            logMessage(LOG_ERROR, "ERROR", "TableSet allocation failed");
            goto folder_cleanup;
        }

        // --- 5. 處理資料 ---
        for (int j = 0; j < total_o_files; j++) {
            char full_path[512];
            snprintf(full_path, sizeof(full_path), "%s\\%s", folder, myFiles[j]);
            logMessage(LOG_INFO, "INFO", "Processing file [%d/%d]: %s", j + 1, total_o_files, myFiles[j]);
            if (loadSingleFileData(full_path, &(myData->tables[j]), 
                                   file_info.particles, file_info.tally_count, 
                                   file_info.label_count, file_info.labels) != 0) {
                logMessage(LOG_ERROR, "ERROR", "File load failed: %s", myFiles[j]);
            }
        }

        // --- 6. 匯出 ---
        if (exportFinalReport(folder, &origin, myData, myFiles) == 0) {
            logMessage(LOG_SUCCESS, "SUCCESS", "Folder processed: %s", folder);
            total_processed_files += total_o_files;
        }

        // --- 7. 統一清理區域 (The Cleanup Label) ---
        folder_cleanup:
        if (dir) closedir(dir);
        if (myFiles) { cleanup_list(myFiles, total_o_files); myFiles = NULL; }
        if (myData) { freeTableSet(myData); myData = NULL; }
        free_file_info(&file_info);
        clearTable(&origin);

        ULONGLONG folder_end = GetTickCount();
        logMessage(LOG_SYSTEM, "SYSTEM", "Folder time: %.2f s", (folder_end - folder_start) / 1000.0);
    }

    if (indices) free(indices);

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
	    
	    // 1. 執行 PowerShell 腳本
	    ShellExecute(NULL, "open", "powershell.exe",
	                 "-ExecutionPolicy Bypass -File merge.ps1",
	                 NULL, SW_HIDE);
	
	    // 2. 顯示第一個視窗 (普通的確定視窗)
	    MessageBox(hwnd,
	               "Excel 整合指令已送出，請檢查 MCNP_Total_Summary.xlsx。",
	               "完成",
	               MB_OK | MB_ICONINFORMATION);
	
	    // 3. 當上面的視窗關閉後，緊接著顯示「畢業視窗」
	    int gradResult = ShowGraduationDialog(hwnd);
	
	    // 根據使用者的選擇紀錄日誌
	    if (gradResult == IDCANCEL) {
	        // 使用者點了「我要延畢...」
	        logMessage(LOG_WARN, "SYSTEM", "使用者選擇延畢，實驗室的燈火將為你而留。");
	        MessageBox(hwnd, "這勇氣令人佩服，實驗室永遠歡迎你。", "開發者的關懷", MB_OK);
	    } else {
	        // 使用者點了「萬歲！」
	        logMessage(LOG_SUCCESS, "SYSTEM", "使用者順利畢業，恭喜脫離苦海！");
	        MessageBox(hwnd, "ya~( °▽°)/自由了\n\n程式已正常結束。\nreturn 0;", "wow 沒有閃退欸", MB_OK);
	    }
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
					" ├──xxx.exe\n"
					" ├──merge.ps1\n"
					" ├──(source).csv\n"
					" ├──(files)\n"
					" │        └──xxx.o\n"
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

// --- 主程式 ---
int wmain() { // 使用 wmain 支援 Unicode 命令列
    AllocConsole();
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);

    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    freopen("CONOUT$", "w", stdout);

    SetConsoleTitleW(L"卷卷又糖糖");

    // 神獸守護畫面
    setColor(LOG_SUCCESS);
    printf("                /\\_/\\      ___      /\\_/\\\n");
    printf("               ( o.o )   /   \\    ( o.o )\n");
    printf("                > ^ <   / /^\\ \\    > ^ <\n");
    printf("             ___/     \\_/ /_\\ \\_/     \\___\n");
    printf("                    神獸守護，程式不閃退!\n");
    setColor(LOG_NORMAL);

    // 檢查預設檔案
    FILE* f = _wfopen(L"source.csv", L"r");
    if (f) {
        logMessage(LOG_INFO, "INFO", "Default 'source.csv' detected.");
        wcscpy(G_SelectedSourcePath, L"source.csv");
        fclose(f);
    } else {
        logMessage(LOG_WARN, "WARN", "'source.csv' missing. Requesting user input...");
        MessageBoxW(NULL, L"偵測不到 source.csv 檔案！\n請點擊確定後手動選取來源 CSV。", L"找不到檔案", MB_OK | MB_ICONINFORMATION);

        if (!SelectSourceFile(NULL, G_SelectedSourcePath)) {
            logMessage(LOG_ERROR, "FATAL", "User cancelled selection. Exiting.");
            MessageBoxW(NULL, L"未選取來源檔案，程式將結束。", L"Error", MB_OK | MB_ICONERROR);
            return 0;
        }
    }

    logMessage(LOG_SUCCESS, "READY", "Current source: %ls", G_SelectedSourcePath);
    printf(" Waiting for user interaction...\n\n");

    // 註冊視窗類別與建立視窗
    HINSTANCE hInstance = GetModuleHandle(NULL);
    WNDCLASS wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = "MCNPDevToolkit";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClass(&wc);

    HWND hwnd = CreateWindow("MCNPDevToolkit", "MCNP 數據整合工具", 
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


