/* ================= KikKoh @2026 =================
   ================ 最終版：支援複選資料夾批次處理 ============ */
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <shellapi.h>
#include <dirent.h>
#include <commdlg.h>

#include "config.h"
#include "tool.h"

#define ID_LISTBOX  101
#define ID_BTN_OK   102
#define ID_BTN_HELP 103

// --- 全域變數 ---
HWND hListBox, hBtnOk;
HFONT hFont;
char G_SelectedSourcePath[MAX_PATH] = {0};

// --- 視窗輔助函式 ---
void SetDefaultFont(HWND hwnd) {
    hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                       VARIABLE_PITCH | FF_SWISS, "Microsoft JhengHei");
    SendMessage(hwnd, WM_SETFONT, (WPARAM)hFont, TRUE);
}

// --- 檔案選取函式 ---
int SelectSourceFile(HWND hwnd, char* outPath) {
    OPENFILENAME ofn;
    char szFile[260] = {0};
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;
    if (GetOpenFileName(&ofn)) {
        strcpy(outPath, ofn.lpstrFile);
        return 1;
    }
    return 0;
}

// --- 核心處理邏輯 ---
void ExecuteBatchProcess(HWND hwnd) {
    int count = SendMessage(hListBox, LB_GETSELCOUNT, 0, 0);
    if (count <= 0) {
        MessageBox(hwnd, "請至少選擇一個資料夾！", "提示", MB_OK | MB_ICONWARNING);
        return;
    }

    int *indices = (int *)malloc(sizeof(int) * count);
    if (indices == NULL) {
        printf("\n [ ERROR ] Memory allocation failed!\n");
        return;
    }
    SendMessage(hListBox, LB_GETSELITEMS, (WPARAM)count, (LPARAM)indices);

    printf("\n------------------------------------------------------------\n");
    printf(" [ BATCH ] Starting batch process for %d directories...\n", count);
    printf("------------------------------------------------------------\n");

    for (int i = 0; i < count; i++) {
        char folder[MAX_PATH];
        SendMessage(hListBox, LB_GETTEXT, indices[i], (LPARAM)folder);
        
        printf(" [%d/%d] Processing: %-30s", i + 1, count, folder);

        int rows;
        DynamicTable origin;
        int total_o_files = 0;
        FileInfo file_info;
        TableSet* myData = NULL;
        char **myFiles = NULL;

        rows = loadSourceCSV(G_SelectedSourcePath, &origin);
        if (rows <= 0) {
            printf("\n  >> [ SKIP ] Unable to read source file: %s\n", G_SelectedSourcePath);
            continue; 
        }

        DIR *dir = opendir(folder);
        struct dirent *ent;
        if (!dir) {
            printf("\n  >> [ SKIP ] Path not found: %s\n", folder);
            clearTable(&origin);
            continue;
        }

        while ((ent = readdir(dir)) != NULL) {
            char *ext = strrchr(ent->d_name, '.');
            if (ext && strcmp(ext, ".o") == 0) {
                total_o_files++;
                if (total_o_files == 1) {
                    char full_path[512];
                    sprintf(full_path, "%s/%s", folder, ent->d_name);
                    file_info = get_file_info(full_path);
                }
            }
        }
        closedir(dir);

        if (total_o_files > 0 && file_info.particles > 0) {
            total_o_files = get_sorted_o_files(folder, &myFiles);
            myData = createTableSet(total_o_files, rows - 2, file_info.label_count + 1);
            
            for (int j = 0; j < total_o_files; j++) {
                char full_path[512];
                snprintf(full_path, sizeof(full_path), "%s/%s", folder, myFiles[j]);
                loadSingleFileData(full_path, &(myData->tables[j]), file_info.particles, 
                                   file_info.tally_count, file_info.label_count, file_info.labels);
            }

            exportFinalReport(folder, &origin, myData, myFiles);
            printf(" -> [ SUCCESS ]\n");
        } else {
            printf(" -> [ EMPTY ]\n");
        }

        clearTable(&origin);
        if (myFiles) cleanup_list(myFiles, total_o_files);
        if (myData) freeTableSet(myData);
    }

    free(indices);

    printf("------------------------------------------------------------\n");
    printf(" [  DONE  ] All folders processed successfully!\n");
    printf("------------------------------------------------------------\n\n");

    int msgboxID = MessageBox(
        hwnd,
        "批次 CSV 處理完成！\n是否要立即執行 PowerShell 整合為分頁 Excel？",
        "批次處理完成",
        MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON1
    );

    if (msgboxID == IDYES) {
        printf(" [ SYSTEM ] Launching Excel integration script...\n");
        ShellExecute(NULL, "open", "powershell.exe", 
                     "-ExecutionPolicy Bypass -File merge.ps1", 
                     NULL, SW_HIDE);
        
        MessageBox(hwnd, "Excel 整合指令已送出，請檢查 Total_Summary_Styled.xlsx。", "完成", MB_OK | MB_ICONINFORMATION);
    }
}

// --- 視窗回呼函式 ---
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
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
					" ├──merge.ps1\n" //            # 設定檔
					" ├──(source.csv)\n"
					" ├──(files)\n"
					" │       └──OO.o\n"
					" ├──(source.csv)\n"
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
    freopen("CONOUT$", "w", stdout);
    //SetConsoleTitle("MRCP Data Toolkit - System Console");
    SetConsoleTitle("卷卷又糖糖");
 
    // 移除低調化
    printf(" __________________________________________________________ \n");
    printf("|                                                          |\n");
    printf("|        M C N P   D A T A   T O O L K I T   v2.5          |\n");
    printf("|__________________________________________________________|\n\n");

    FILE *f = fopen("source.csv", "r");
    if (f) {
        printf(" [ INFO ] Default 'source.csv' detected.\n");
        strcpy(G_SelectedSourcePath, "source.csv");
        fclose(f);
    } else {
        printf(" [ WARN ] 'source.csv' missing. Requesting user input...\n");
        MessageBox(NULL, "偵測不到 source.csv 檔案！\n請點擊確定後手動選取來源 CSV。", "找不到檔案", MB_OK | MB_ICONINFORMATION);
        
        if (!SelectSourceFile(NULL, G_SelectedSourcePath)) {
            printf(" [ FATAL ] User cancelled selection. Exiting.\n");
            MessageBox(NULL, "未選取來源檔案，程式將結束。", "Error", MB_OK | MB_ICONERROR);
            return 0;
        }
    }

    printf(" [ READY ] Current source: %s\n\n", G_SelectedSourcePath);
    printf(" Waiting for user interaction...\n");

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

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!IsDialogMessage(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // --- 結束前的倒數邏輯 ---
    printf("\n\n============================================================\n");
    printf(" 感謝您的使用。\n");
    for(int i = 3; i > 0; i--) {
        printf(" 程式將在 %d 秒後自動關閉...\r", i);
        fflush(stdout);
        Sleep(1000);
    }
    printf("\n============================================================\n");

    FreeConsole();
    return 0;
}
