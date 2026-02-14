/* ================= KikKoh @2026 =================
================ 版權所有，翻印必究 =============== */

#include <wininet.h>
#include <urlmon.h>

#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "urlmon.lib")
#pragma comment(lib, "shell32.lib")

// ===== 定義目標 =====
// 權限盾牌 
#ifndef MB_ICONSHIELD
#define MB_ICONSHIELD 0x0000004CL
#endif
// 版本目標
#define CURRENT_VERSION "v2.7.1"
#define VERSION_URL "https://raw.githubusercontent.com/10809104/MCNP-Data-Toolkit/main/version.txt"

/* ================= 資料結構 ================= */
// 定義資料型別枚舉，用來標記 Cell 目前儲存的是哪種資料
typedef enum {
    TYPE_STRING, 
    TYPE_DOUBLE
} DataType;

// 定義單一單元格結構
typedef struct {
    DataType type;      // 記錄型別標籤 (Tag)
    union {
        char* s;        // 儲存字串指標 (指向動態分配的記憶體)
        double d;       // 儲存倍精度浮點數
    } data;             // union 空間共用，節省記憶體
} Cell;

// 定義動態二維表格結構
typedef struct {
    int rows;           // 表格列數
    int cols;           // 表格欄數
    Cell** table;       // 二級指標：table[row][col]
} DynamicTable;

// 定義表格集合結構，用來管理多個 .o 檔的數據
typedef struct {
    int file_count;       // 檔案總數
    DynamicTable* tables; // 指向 DynamicTable 陣列的指標
} TableSet;

// 定義單一 .o 檔得到的固定數據
typedef struct {
    long long particles; // 歷史粒子數
    char **labels;       // 動態二維陣列
    int label_count;     // 標籤數量
    int tally_count;     // 一行有多少數據 
    int total_count;     // 總共有多少數據 
} FileInfo;

/* =================  initial functions  ================= */
// --- 函數實作 ---

/**
 * 初始化單張表格
 * @param dt 指向表格結構的指標
 * @param r  列數
 * @param c  欄數
 */
void initTable(DynamicTable* dt, int r, int c) {
    dt->rows = r;
    dt->cols = c;
    
    // 第一層：配置指向各列指標的空間 (指標陣列)
    dt->table = (Cell**)malloc(r * sizeof(Cell*));
    
    for (int i = 0; i < r; i++) {
        // 第二層：為每一列配置實際存放 Cell 的空間
        dt->table[i] = (Cell*)malloc(c * sizeof(Cell));
        
        // 初始化每個 Cell 的預設值，避免產生隨機垃圾值
        for (int j = 0; j < c; j++) {
            dt->table[i][j].type = TYPE_DOUBLE;
            dt->table[i][j].data.d = 0.0;
        }
    }
}

/**
 * 釋放單張表格的記憶體 (由內而外釋放)
 */
void clearTable(DynamicTable* dt) {
    if (!dt || !dt->table) return;

    for (int i = 0; i < dt->rows; i++) {
        for (int j = 0; j < dt->cols; j++) {
            // 重要：如果單元格存的是動態分配的字串，必須先 free 字串
            if (dt->table[i][j].type == TYPE_STRING && dt->table[i][j].data.s != NULL) {
                free(dt->table[i][j].data.s);
            }
        }
        // 釋放每一列的記憶體
        free(dt->table[i]);
    }
    // 釋放指標陣列本身
    free(dt->table);
    dt->table = NULL; // 設為 NULL 避免野指標 (dangling pointer)
}

/**
 * 建立並初始化 TableSet (管理多個檔案)
 * r, c 是預期每個 .o 檔會有的行列數
 */
TableSet* createTableSet(int num_files, int r, int c) {
    if (num_files <= 0) return NULL;

    // 1. 配置管理結構
    TableSet* ts = (TableSet*)malloc(sizeof(TableSet));
    if (!ts) return NULL;

    ts->file_count = num_files;
    
    // 2. 配置 Table 陣列空間
    ts->tables = (DynamicTable*)malloc(num_files * sizeof(DynamicTable));
    if (!ts->tables) {
        free(ts);
        return NULL;
    }

    // 3. 逐一初始化子表格
    for (int i = 0; i < num_files; i++) {
        // 這裡調用你之前寫好的 initTable
        initTable(&(ts->tables[i]), r, c);
        
        // 額外保險：如果 initTable 失敗（內部 malloc 失敗），
        // 實務上這裡可能需要更複雜的清理邏輯，但基本使用先這樣即可
    }
    
    return ts;
}

/**
 * 釋放整個 TableSet 資源
 */
void freeTableSet(TableSet* ts) {
    if (!ts) return;
    
    for (int i = 0; i < ts->file_count; i++) {
        // 呼叫單張表格的釋放函數
        clearTable(&(ts->tables[i]));
    }
    // 釋放表格陣列與管理結構
    free(ts->tables);
    free(ts);
}

/**
 * 判斷是否為管理員
 */
BOOL IsRunningAsAdmin()
{
    BOOL isAdmin = FALSE;
    PSID adminGroup = NULL;

    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;

    if (AllocateAndInitializeSid(&NtAuthority,
                                 2,
                                 SECURITY_BUILTIN_DOMAIN_RID,
                                 DOMAIN_ALIAS_RID_ADMINS,
                                 0,0,0,0,0,0,
                                 &adminGroup))
    {
        CheckTokenMembership(NULL, adminGroup, &isAdmin);
        FreeSid(adminGroup);
    }

    return isAdmin;
}

/**
 * 版本更新
 */
void PerformUpdate(const char* downloadUrl) 
{ 
    char currentExe[MAX_PATH] = {0}; 
    char tempPath[MAX_PATH] = {0}; 
    char newExePath[MAX_PATH] = {0}; 
    char batPath[MAX_PATH] = {0}; 

    GetModuleFileName(NULL, currentExe, MAX_PATH); 
    
    printf("[UPDATE] Downloading new version... Please wait.\n");
    // --- 1. 權限檢查 ---
    // 嘗試在 exe 所在目錄建立測試檔案，判斷是否有寫入權限
    char testFilePath[MAX_PATH];
    sprintf(testFilePath, "%s.tmp", currentExe);
    FILE *testFile = fopen(testFilePath, "w");
    if (!testFile) {
        int result = MessageBox(NULL, "更新需要管理員權限，是否以管理員身份重新啟動並更新？", "權限不足", MB_YESNO | MB_ICONSHIELD);
        if (result == IDYES) {
            // 使用 "runas" 觸發 UAC 提升權限重新啟動自己
            if ((INT_PTR)ShellExecute(NULL, "runas", currentExe, NULL, NULL, SW_SHOWNORMAL) > 32) {
                ExitProcess(0);
            }
        }
        return; // 使用者選否或啟動失敗則取消更新
    } else {
        fclose(testFile);
        remove(testFilePath); // 測試成功後刪除
    }

    // --- 2. 下載流程 ---
    GetTempPath(MAX_PATH, tempPath); 
    sprintf(newExePath, "%sMRCP_Update_Temp.exe", tempPath); 
    sprintf(batPath, "%supdate_mrcp.bat", tempPath); 

    printf("[UPDATE] Downloading new version...\n"); 
    HRESULT hr = URLDownloadToFile(NULL, downloadUrl, newExePath, 0, NULL); 

    if (FAILED(hr)) { 
        MessageBox(NULL, "下載失敗，請檢查網路連線或手動更新。", "錯誤", MB_ICONERROR); 
        return; 
    } 

    // --- 3. 建立自毀式更新腳本 ---
    FILE* f = fopen(batPath, "w"); 
    if (!f) { 
        MessageBox(NULL, "無法建立臨時指令腳本。", "錯誤", MB_ICONERROR); 
        return; 
    } 

    // 腳本邏輯：無限循環直到舊版被刪除 -> 移動新版覆蓋舊版 -> 啟動新版 -> 刪除腳本自己
    fprintf(f, 
        "@echo off\n" 
        "title MRCP Updater\n"
        "echo 正在等待程式關閉並套用更新...\n" 
        ":loop\n" 
        "del \"%s\" >nul 2>&1\n" 
        "if exist \"%s\" (\n" 
        "  timeout /t 1 /nobreak >nul\n" 
        "  goto loop\n" 
        ")\n" 
        "move /Y \"%s\" \"%s\" >nul\n" 
        "start \"\" \"%s\"\n" 
        "del \"%%~f0\" & exit\n", // 這裡使用 & exit 確保指令執行完立即關閉
        currentExe, 
        currentExe, 
        newExePath, 
        currentExe, 
        currentExe 
    ); 
    fclose(f); 

    // --- 4. 執行更新 ---
    // 使用 SW_HIDE 讓使用者看不到黑色視窗閃過
    ShellExecute(NULL, "open", batPath, NULL, NULL, SW_HIDE); 

    ExitProcess(0); 
}

/**
 * 版本檢查
 */
/**
 * 版本檢查 (語義化版本比對版)
 */
void CheckForUpdates()
{
    printf("[SYSTEM] Checking for updates...\n");

    HINTERNET hInternet = InternetOpen("MRCP_Tool",
                                       INTERNET_OPEN_TYPE_DIRECT,
                                       NULL, NULL, 0);

    if (!hInternet) {
        printf("InternetOpen failed.\n");
        return;
    }

    HINTERNET hConnect = InternetOpenUrl(hInternet,
                                          VERSION_URL,
                                          NULL,
                                          0,
                                          INTERNET_FLAG_RELOAD,
                                          0);

    if (!hConnect) {
        printf("Unable to connect to version server.\n");
        InternetCloseHandle(hInternet);
        return;
    }

    char remoteVersion[32] = {0};
    DWORD bytesRead = 0;

    if (!InternetReadFile(hConnect,
                          remoteVersion,
                          sizeof(remoteVersion) - 1,
                          &bytesRead))
    {
        printf("Failed to read version info.\n");
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return;
    }

    remoteVersion[bytesRead] = '\0';
    remoteVersion[strcspn(remoteVersion, "\r\n")] = 0; // 去除換行

    if (strlen(remoteVersion) == 0) {
        printf("Invalid version data.\n");
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);
        return;
    }

    printf("Remote version: %s\n", remoteVersion);
    printf("Current version: %s\n", CURRENT_VERSION);

    // --- 新的版本比較邏輯 ---
    int r_major = 0, r_minor = 0, r_patch = 0; // 預設為 0
    int c_major = 0, c_minor = 0, c_patch = 0;

    // 解析字串 (假設格式為 v2.7.1 或 2.7.1)
    // %*[^0-9] 意思是跳過開頭所有非數字的字元(例如 'v')
    int r_count = sscanf(remoteVersion, "%*[^0-9]%d.%d.%d", &r_major, &r_minor, &r_patch);
    int c_count = sscanf(CURRENT_VERSION, "%*[^0-9]%d.%d.%d", &c_major, &c_minor, &c_patch);

    BOOL hasUpdate = FALSE;
    if (r_count >= 2 && c_count >= 2) { // 至少要有 Major.Minor 才能比較
        if (r_major > c_major) {
            hasUpdate = TRUE;
        } else if (r_major == c_major) {
            if (r_minor > c_minor) {
                hasUpdate = TRUE;
            } else if (r_minor == c_minor) {
                if (r_patch > c_patch) {
                    hasUpdate = TRUE;
                }
            }
        }
    }

    if (hasUpdate)
    {
        char msg[256];
        sprintf(msg,
                "偵測到新版本：%s\n目前版本：%s\n是否立即下載並更新？",
                remoteVersion,
                CURRENT_VERSION);

        int result = MessageBox(NULL,
                                msg,
                                "發現更新",
                                MB_YESNO | MB_ICONQUESTION | MB_TOPMOST);

        if (result == IDYES)
        {
            const char* exeUrl =
                "https://github.com/10809104/MCNP-Data-Toolkit/releases/latest/download/MRCP_Tool.exe";

            PerformUpdate(exeUrl);
        }
    }
    else
    {
        printf("[SYSTEM] You are using the latest version.\n");
    }

    InternetCloseHandle(hConnect);
    InternetCloseHandle(hInternet);
}
