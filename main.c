/* ================= KikKoh @2026 =================
   ================ 最終版：支援複選資料夾批次處理 ============ */

// 關鍵：不再全域定義 UNICODE，改由顯式呼叫 W 版本 API 解決衝突
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <shellapi.h>
#include <dirent.h>
#include <commdlg.h>
#include <time.h>
#include <stdarg.h>

#include "config.h"
#include "tools.h"

#define MAX_PATH 260

// 定義 UI 元件 ID
#define ID_LISTBOX  101
#define ID_BTN_OK   102
#define ID_BTN_HELP 103

// ===== 顏色定義 =====
#define LOG_INFO    11 // 淺藍
#define LOG_SUCCESS 10 // 綠色
#define LOG_WARN    14 // 黃色
#define LOG_ERROR   12 // 紅色
#define LOG_SYSTEM  9  // 藍色
#define LOG_NORMAL  7  // 預設白色

// --- 全域變數 ---
HWND hListBox, hBtnOk;
HFONT hFont;
// 這裡維持 char，因為你的 loadSourceCSV 內部使用的是 char*
char G_SelectedSourcePath[MAX_PATH] = {0}; 
HANDLE hConsole;

/** 變更 Console 輸出顏色 */
void setColor(int color) {
    SetConsoleTextAttribute(hConsole, color);
}

/** 格式化輸出日誌 */
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
/** 設定視窗元件為微軟正黑體 */
void SetDefaultFont(HWND hwnd) {
    hFont = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                       VARIABLE_PITCH | FF_SWISS, L"Microsoft JhengHei");
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

/** 開啟檔案選取對話框 (核心轉接邏輯) */
int SelectSourceFile(HWND hwnd, char* outPath) {
    OPENFILENAMEW ofn;
    WCHAR szFile[MAX_PATH] = {0};
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetOpenFileNameW(&ofn)) {
        // 取得 Unicode 路徑後，轉回 char 給你的後續函式使用
        WideCharToMultiByte(CP_ACP, 0, ofn.lpstrFile, -1, outPath, MAX_PATH, NULL, NULL);
        return 1;
    }
    return 0;
}

// --- 自定義按鈕文字的 Hook 邏輯 ---
LRESULT CALLBACK MsgBoxHookProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode == HCBT_ACTIVATE) {
        HWND hwndMsgBox = (HWND)wParam;
        SetDlgItemTextW(hwndMsgBox, IDOK, L"萬歲！");
        SetDlgItemTextW(hwndMsgBox, IDCANCEL, L"我要延畢...");
        return 0;
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

/** 顯示畢業視窗 */
int ShowGraduationDialog(HWND hwnd) {
    HHOOK hHook = SetWindowsHookExW(WH_CBT, MsgBoxHookProc, NULL, GetCurrentThreadId());
    int result = MessageBoxW(hwnd, L"恭喜操作完成，可以順利畢業了！", L"順利執行完畢", MB_OKCANCEL | MB_ICONINFORMATION);
    UnhookWindowsHookEx(hHook);
    return result;
}

// ===== 核心批次處理邏輯 =====
void ExecuteBatchProcess(HWND hwnd) {
    ULONGLONG total_start = GetTickCount64();
    int total_processed_files = 0;
    int total_directories = 0;

    int count = SendMessage(hListBox, LB_GETSELCOUNT, 0, 0);
    if (count <= 0) {
        MessageBoxW(hwnd, L"請至少選擇一個資料夾！", L"提示", MB_OK | MB_ICONWARNING);
        return;
    }

    int *indices = (int *)malloc(sizeof(int) * count);
    if (!indices) return;

    SendMessage(hListBox, LB_GETSELITEMS, (WPARAM)count, (LPARAM)indices);
    logMessage(LOG_SYSTEM, "SYSTEM", "Batch process started (%d directories)", count);

    for (int i = 0; i < count; i++) {
        ULONGLONG folder_start = GetTickCount64();
        total_directories++;

        DynamicTable origin = {0}; 
        FileInfo file_info = {0, NULL, 0, 0, 0};
        TableSet* myData = NULL;
        char **myFiles = NULL;
        int total_o_files = 0;
        DIR *dir = NULL;

        char folder[MAX_PATH];
        SendMessage(hListBox, LB_GETTEXT, indices[i], (LPARAM)folder);
        logMessage(LOG_INFO, "INFO", "[%d/%d] Processing directory: %s", i + 1, count, folder);

        if (loadSourceCSV(G_SelectedSourcePath, &origin) <= 0) {
            logMessage(LOG_WARN, "WARNING", "Failed to read source CSV: %s", G_SelectedSourcePath);
            goto folder_cleanup;
        }

        dir = opendir(folder);
        if (!dir) goto folder_cleanup;

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
        closedir(dir); dir = NULL;

        if (total_o_files <= 0 || file_info.total_count <= 0) goto folder_cleanup;

        total_o_files = get_sorted_o_files(folder, &myFiles);
        myData = createTableSet(total_o_files, file_info.total_count + 1, file_info.label_count + 1);
        if (!myData) goto folder_cleanup;

        for (int j = 0; j < total_o_files; j++) {
            char full_path[512];
            snprintf(full_path, sizeof(full_path), "%s\\%s", folder, myFiles[j]);
            loadSingleFileData(full_path, &(myData->tables[j]), file_info.particles, file_info.tally_count, file_info.label_count, file_info.labels);
        }

        if (exportFinalReport(folder, &origin, myData, myFiles) == 0) {
            logMessage(LOG_SUCCESS, "SUCCESS", "Folder processed: %s", folder);
            total_processed_files += total_o_files;
        }

        folder_cleanup:
        if (dir) closedir(dir);
        if (myFiles) { cleanup_list(myFiles, total_o_files); }
        if (myData) { freeTableSet(myData); }
        free_file_info(&file_info);
        clearTable(&origin);
        logMessage(LOG_SYSTEM, "SYSTEM", "Folder time: %.2f s", (GetTickCount64() - folder_start) / 1000.0);
    }

    free(indices);
    logMessage(LOG_SYSTEM, "SUMMARY", "Total directories : %d", total_directories);
    
    int msgboxID = MessageBoxW(hwnd, L"批次處理完成！\n是否要整合為分頁 Excel？", L"完成", MB_YESNO | MB_ICONQUESTION);
    if (msgboxID == IDYES) {
        ShellExecuteW(NULL, L"open", L"powershell.exe", L"-ExecutionPolicy Bypass -File merge.ps1", NULL, SW_HIDE);
        MessageBoxW(hwnd, L"Excel 整合指令已送出。", L"完成", MB_OK);
        
        int gradResult = ShowGraduationDialog(hwnd);
        if (gradResult == IDCANCEL) {
            logMessage(LOG_WARN, "SYSTEM", "使用者選擇延畢。");
        } else {
            logMessage(LOG_SUCCESS, "SYSTEM", "使用者順利畢業！");
        }
    }
}

// --- 視窗回呼函式 (WndProc) ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            // 這裡強烈建議使用 W 版本 API 建立 UI
            HWND hGroup = CreateWindowExW(0, L"BUTTON", L" 目錄複選 (Ctrl/Shift) ", WS_VISIBLE | WS_CHILD | BS_GROUPBOX, 10, 10, 265, 180, hwnd, NULL, NULL, NULL);
            SetDefaultFont(hGroup);

            hListBox = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", NULL, WS_VISIBLE | WS_CHILD | LBS_STANDARD | WS_VSCROLL | LBS_EXTENDEDSEL, 25, 40, 235, 130, hwnd, (HMENU)ID_LISTBOX, NULL, NULL);
            SetDefaultFont(hListBox);

            HWND hBtnHelp = CreateWindowExW(0, L"BUTTON", L"說明 (Help)", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 15, 205, 80, 35, hwnd, (HMENU)ID_BTN_HELP, NULL, NULL);
            SetDefaultFont(hBtnHelp);

            hBtnOk = CreateWindowExW(0, L"BUTTON", L"批次處理", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 155, 205, 120, 35, hwnd, (HMENU)ID_BTN_OK, NULL, NULL);
            SetDefaultFont(hBtnOk);

            WIN32_FIND_DATAW findData;
            HANDLE hFind = FindFirstFileW(L".\\*", &findData);
            if (hFind != INVALID_HANDLE_VALUE) {
                do {
                    if ((findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && findData.cFileName[0] != '.') {
                        SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)findData.cFileName);
                    }
                } while (FindNextFileW(hFind, &findData));
                FindClose(hFind);
            }
            break;
        }
        case WM_COMMAND:
            if (LOWORD(wParam) == ID_BTN_OK) ExecuteBatchProcess(hwnd);
            if (LOWORD(wParam) == ID_BTN_HELP) {
                MessageBoxW(hwnd, L"使用說明：\n1. 選取 Source CSV\n2. 選取資料夾\n3. 點擊批次處理", L"說明", MB_OK | MB_ICONINFORMATION);
            }
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// --- 主程式進入點 ---
int main() {
    AllocConsole();
    // 設定輸出編碼為 UTF-8 以支援神獸 ASCII Art
    SetConsoleOutputCP(65001);
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    freopen("CONOUT$", "w", stdout);

    SetConsoleTitleW(L"卷卷又糖糖 - 神獸守護版");

    setColor(LOG_SUCCESS);
    printf("                /\\_/\\      ___      /\\_/\\\n");
    printf("               ( o.o )   /   \\    ( o.o )\n");
    printf("                > ^ <   / /^\\ \\    > ^ <\n");
    printf("             ___/     \\_/ /_\\ \\_/     \\___\n");
    printf("                    神獸守護，程式不閃退!\n");
    setColor(LOG_NORMAL);

    // 檢查預設檔案
    if (_waccess(L"source.csv", 0) == 0) {
        logMessage(LOG_INFO, "INFO", "Default 'source.csv' detected.");
        strcpy(G_SelectedSourcePath, "source.csv");
    } else {
        logMessage(LOG_WARN, "WARN", "Requesting user input...");
        if (!SelectSourceFile(NULL, G_SelectedSourcePath)) return 0;
    }

    HINSTANCE hInst = GetModuleHandle(NULL);
    WNDCLASSW wc = {0};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"MCNPDevToolkitClass"; // 寬字元類別名稱
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(0, L"MCNPDevToolkitClass", L"MCNP 數據整合工具", 
                               WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_VISIBLE, 
                               CW_USEDEFAULT, CW_USEDEFAULT, 300, 320, NULL, NULL, hInst, NULL);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        if (!IsDialogMessageW(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    printf("\n 程式結束，3 秒後關閉...\n");
    Sleep(3000);
    return 0;
}
