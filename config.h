/* ================= KikKoh @2026 =================
================ 版權所有，翻印必究 =============== */

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
