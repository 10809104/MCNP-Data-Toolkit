/* ================= KikKoh @2026 =================
================ 版權所有，翻印必究 =============== */
#define LINE_BUF 1024
#define SPACE_THRESHOLD 5
#define TALLY_KEY_LEN 11
#define TALLY_KEY "   tally   "

/* =================  Parser functions  ================= */

/**
 * @brief 讀取 source.csv 並將前兩欄存入 DynamicTable
 * @details 此函數執行兩次掃描：第一次計算總行數以分配記憶體，第二次解析 CSV 內容。
 * 解析時具備簡單的狀態機處理引號，並自動移除數值中的千分位逗號。
 * @param filename 檔案路徑與名稱
 * @param dt       指向預先宣告的 DynamicTable 結構指標
 * @return int     成功返回總列數，失敗返回 -1
 */
int loadSourceCSV(const char *filename, DynamicTable *dt) {
    char line[LINE_BUF + 1];
    int totalRows = 0;

    FILE *fp = fopen(filename, "r");
    if (!fp) {
        printf("錯誤：無法開啟 %s\n", filename);
        return -1;
    }

    // --- 第一次掃描：統計有效行數 ---
    while (fgets(line, sizeof(line), fp)) {
        if (strlen(line) > 1) totalRows++; // 略過完全空白的行
    }
    rewind(fp); // 將檔案指標重置回開頭

    // 根據掃描結果初始化表格記憶體 (固定讀取前 2 欄)
    initTable(dt, totalRows, 2);

    int r = 0;
    while (fgets(line, sizeof(line), fp) && r < totalRows) {
        line[strcspn(line, "\r\n")] = 0; // 移除行尾的換行符號

        char *ptr = line;
        for (int col = 0; col < 2; col++) {
            char temp[LINE_BUF] = {0};
            int t_idx = 0;
            int in_quotes = 0;

            // --- 狀態機解析單個 CSV 欄位 ---
            while (*ptr != '\0') {
                if (*ptr == '"') {
                    in_quotes = !in_quotes; // 切換「引號內」模式
                } else if (*ptr == ',' && !in_quotes) {
                    ptr++; // 遇到作為分隔符的逗號，跳過並結束此欄位讀取
                    break; 
                } else {
                    // 若非逗號，或是在引號內的逗號（如 "11,513"）
                    // 根據需求：我們選擇忽略數字中的逗號以方便後續轉為數值
                    if (*ptr != ',') {
                        if (t_idx < sizeof(temp) - 1) temp[t_idx++] = *ptr;
                    } 
                }
                ptr++;
            }

            // 將解析結果存入表格結構
            dt->table[r][col].type = TYPE_STRING;
            dt->table[r][col].data.s = strdup(temp);
            // 確認 strdup 是否成
            if (dt->table[r][col].data.s == NULL && strlen(temp) > 0) {
                fprintf(stderr, "記憶體分配失敗 (strdup)\n");
                fclose(fp);
                clearTable(dt);
                return -1;
            }
        }
        r++;
    }

    fclose(fp);
    return totalRows;
}

/**
 * @brief 釋放 FileInfo 結構內部的動態記憶體
 */
void free_file_info(FileInfo *info) {
    if (info && info->labels) {
        for (int k = 0; k < info->label_count; k++) {
            if (info->labels[k]) free(info->labels[k]);
        }
        free(info->labels);
        info->labels = NULL; // 避免 Double Free
    }
}

/**
 * @brief 整合型副程式：抓取 MCNP 輸出檔的標籤名稱與總粒子數
 * @details 包含三個步驟：
 * 1. 偵測 Tally 數量與原始標籤。
 * 2. 讀取首行數據進行「校準」，若某標籤下無數據則自動剔除該標籤並向前遞補。
 * 3. 從檔案末尾反向搜尋執行結束的粒子數。
 * @param path 檔案路徑
 * @return FileInfo 包含標籤清單、標籤數及粒子數的結構
 */
FileInfo get_file_info(const char *path) {
    FileInfo report = {-1, NULL, 0, 0, 0};
    FILE *fp = fopen(path, "r");
    if (!fp) return report;

    char line[LINE_BUF];
    
    // --- 新增部分：全檔掃描統計 "tally" 總出現次數 ---
    while (fgets(line, sizeof(line), fp)) {
        char *ptr = line;
        while ((ptr = strstr(ptr, TALLY_KEY)) != NULL) {
            report.total_count++; // 統計全檔出現關鍵字的總數
            ptr += TALLY_KEY_LEN; // 跳過 "    tally    " 長度繼續搜尋
        }
    }
    // printf("檔案 %s 出現 %d 個數據\n", path, report.total_count)
    rewind(fp); // 統計完畢，將檔案指標重置回開頭以進行後續解析
    
    // --- 第一部分：從檔案中段抓取結構資訊 ---
    while (fgets(line, sizeof(line), fp)) {
        // 尋找關鍵字行，標示數據區塊的開始
        if (strstr(line, TALLY_KEY)) {
            
            // --- 步驟 A: 統計此行出現幾次 "tally"，代表有幾組數據 (Groups) ---
            char *ptr = line;
            while ((ptr = strstr(ptr, TALLY_KEY)) != NULL) {
                report.tally_count++;
                ptr += TALLY_KEY_LEN; // 跳過關鍵字繼續搜尋
            }

            // --- 步驟 B: 解析標籤行 (Labels) ---
            if (fgets(line, sizeof(line), fp)) {
                char temp[LINE_BUF];
                strcpy(temp, line);
                
                // 1. 計算該行總欄位數 (排除 nps 標籤)
                int total_fields = 0;
                char *token = strtok(temp, " \t\n");
                while (token) {
                    if (strcmp(token, "nps") != 0) total_fields++;
                    token = strtok(NULL, " \t\n");
                }
            
                // 計算每一組 Tally 應有的標籤數 (例如 15 欄 / 3 組 = 5 標籤/組)
                if (report.tally_count > 0) {
                    report.label_count = total_fields / report.tally_count;
                }
            
                // 2. 根據每一組的標籤數分配動態記憶體 (字串陣列)
                if (report.label_count > 0) {
                    report.labels = (char **)malloc(report.label_count * sizeof(char *));
                    
                    if (!report.labels) {
                        fprintf(stderr, "無法為標籤分配記憶體\n");
                        fclose(fp);
                        return report;
                    }
                    for (int k = 0; k < report.label_count; k++) report.labels[k] = NULL;
                    
                    // 3. 只填入「第一組」的標籤名稱作為代表
                    strcpy(temp, line);
                    token = strtok(temp, " \t\n");
                    int idx = 0;
					while (token && idx < report.label_count) {
					    if (strcmp(token, "nps") != 0) {
					
					        report.labels[idx] = strdup(token);
					
					        if (!report.labels[idx]) {
					            free_file_info(&report);
					            fclose(fp);
					            return report;
					        }
					
					        idx++;
					    }
					
					    token = strtok(NULL, " \t\n");
					}
                    
                }
            }
            
            // --- 步驟 C: 數據對齊校準 (Calibration) ---
            // 讀取第一筆數據行，確認哪些標籤是有實際對應數據的
            if (fgets(line, sizeof(line), fp)) { 
                char *ptr = line;
                int offset, space_len;
                long long nps_val;
            
                // A. 先跳過開頭的 NPS 數值
                if (sscanf(ptr, "%lld%n", &nps_val, &offset) == 1) {
                    ptr += offset;
                }
            
                // B. 檢查標籤與數據的對應狀況
                int actual_idx = 0; // 記錄實際有數據的標籤索引
                for (int i = 0; i < report.label_count; i++) {
                    // 偵測數據間的空格長度
                    space_len = strspn(ptr, " ");
                    
                    // 若空格過長 (例如 > SPACE_THRESHOLD)，通常代表該輸出位置為空，數據缺失
                    if (space_len > SPACE_THRESHOLD) { 
                        free(report.labels[i]); // 釋放無效標籤
                        report.labels[i] = NULL;
                    } else {
                        // 偵測到數據，讀取並將指標移至下一個數值
                        double temp_val;
                        if (sscanf(ptr, "%lf%n", &temp_val, &offset) == 1) {
                            ptr += offset;
                            
                            // 若前方有標籤被剔除，將後方有效標籤往前遞補
                            if (actual_idx != i) {
                                report.labels[actual_idx] = report.labels[i];
                                report.labels[i] = NULL;
                            }
                            actual_idx++;
                        }
                    }
                }
                // C. 更新最終有效的標籤數量
                report.label_count = actual_idx;
            }
            
            break; // 完成結構抓取，跳出檔案掃描
        }
    }
    
    // --- 第二部分：反向掃描抓取粒子總數 ---
    // 從檔案末尾向前跳 1024 位元組，效率比從頭掃描高
    fseek(fp, 0, SEEK_END);
    long file_size = ftell(fp);
    long read_pos = (file_size > LINE_BUF) ? (file_size - LINE_BUF) : 0;
    fseek(fp, read_pos, SEEK_SET);

    size_t bytes_read = fread(line, 1, LINE_BUF, fp);
    line[bytes_read] = '\0';

    char *target = "run terminated when";
    char *pos = strstr(line, target);
    if (pos) {
        sscanf(pos + strlen(target), "%lld", &report.particles);
    }

    fclose(fp);
    return report;
}

/**
 * @brief 檔名排序用的比較函數 (qsort 呼叫)
 */
int compareNames(const void *a, const void *b) {
    return strcmp(*(const char **)a, *(const char **)b);
}

/**
 * @brief 釋放字串清單佔用的記憶體
 */
void cleanup_list(char **list, int count) {
    if (!list) return;
    for (int i = 0; i < count; i++) {
        free(list[i]);
    }
    free(list);
}

/**
 * @brief 取得指定資料夾內所有 .o 檔案並排序
 * @param path      資料夾路徑
 * @param list_ptr  回傳檔案清單指標的地址
 * @return int      找到的檔案總數，失敗返回 -1
 */
int get_sorted_o_files(const char *path, char ***list_ptr) {
    DIR *dir = opendir(path);
    if (!dir) {
        perror("無法開啟資料夾");
        return -1;
    }

    int count = 0;
    int capacity = 25; // 初始分配空間
    char **fileList = malloc(capacity * sizeof(char *));
    if (!fileList) {
        perror("無法分配檔案列表記憶體");
        closedir(dir);
        return -1;
    }

    struct dirent *ent;
    while ((ent = readdir(dir)) != NULL) {
        char *ext = strrchr(ent->d_name, '.');
        
        // 只篩選字尾為 .o 的檔案
        if (ext && strcmp(ext, ".o") == 0) {
            // 動態擴展陣列容量
            if (count >= capacity) {
                capacity *= 2;
                // --- 記憶體檢查 : realloc 擴展容量 ---
                char **tempList = realloc(fileList, capacity * sizeof(char *));
                if (!tempList) {
                    closedir(dir);
                    cleanup_list(fileList, count);
                    return -1;
                }
                fileList = tempList;
            }
            char *tmp = strdup(ent->d_name);
            if (!tmp) {
			    closedir(dir);
			    cleanup_list(fileList, count);
			    return -1;
			}
			fileList[count++] = tmp;
        }
    }
    closedir(dir);

    // 進行字母順序排序
    if (count > 0) {
        qsort(fileList, count, sizeof(char *), compareNames);
    }

    *list_ptr = fileList; 
    return count;
}

/**
 * @brief 從單一輸出檔載入特定粒子數下的數據
 * @param path      檔案路徑
 * @param dt        目標 DynamicTable
 * @param particles 要搜尋的目標粒子數 (NPS)
 * @param groups    Tally 組數
 * @param set       每組標籤數
 * @param labels    標籤名稱陣列
 * @return int      成功返回 0，失敗返回 -1
 */
int loadSingleFileData(const char* path, DynamicTable* dt, long long particles, int groups, int set, char **labels) {
    FILE *fp = fopen(path, "r");
    if (!fp) return -1;

    int state = 0; // 0: 尋找 Tally 區塊, 1: 尋找數據列
    
    int count = 0; // 記錄處理了幾次大區塊
    char line[LINE_BUF]; 
    
    // --- 設定表格標頭 (Row 0) ---
    dt->table[0][0].type = TYPE_STRING;
    dt->table[0][0].data.s = strdup("no."); 
    if (!dt->table[0][0].data.s) {
        fclose(fp);
        return -1;          // 表頭失敗，直接返回（無需清理其他，因為尚未配置）
    }
    for (int i = 0; i < set; i++) {
        dt->table[0][i+1].type = TYPE_STRING;
        dt->table[0][i+1].data.s = strdup(labels[i]);
        if (!dt->table[0][i+1].data.s) {
            // 釋放之前已成功配置的表頭字串
            for (int j = 0; j < i; j++) {
                free(dt->table[0][j+1].data.s);
                dt->table[0][j+1].data.s = NULL;
            }
            free(dt->table[0][0].data.s);
            dt->table[0][0].data.s = NULL;
            fclose(fp);
            return -1;
        }
    }

    while (fgets(line, sizeof(line), fp)) {
        // 1. 抓取包含 "tally" 的那一列，提取編號 (如 4, 14, 24)
        if (strstr(line, TALLY_KEY) && state == 0) {
            int tally_count = 0;
            char *ptr = line;
            
            while ((ptr = strstr(ptr, TALLY_KEY)) != NULL && tally_count < groups) {
                ptr += TALLY_KEY_LEN; // 跳過關鍵字
                while (*ptr == ' ' || *ptr == '\t') ptr++; // 略過空白
                
                int len = 0;
                while (ptr[len] != ' ' && ptr[len] != '\t' && ptr[len] != '\0' && 
                       ptr[len] != '\n' && ptr[len] != '\r') {
                    len++;
                }

                if (len > 1) {
                    // 計算此數據應存放的列索引 (考慮多個區塊的偏移)
                    int target_row = count * groups + tally_count + 1;
                    if (target_row >= dt->rows) {
					    fclose(fp);
					    return -1;
					}
                    dt->table[target_row][0].type = TYPE_STRING;
                    
                    // 提取 Tally 編號並去尾 (依據原邏輯)
                    char *temp_id = (char *)malloc(len); 
                    if (temp_id) {
                        strncpy(temp_id, ptr, len - 1); 
                        temp_id[len - 1] = '\0';
                        dt->table[target_row][0].data.s = temp_id;
                        // printf("get no. %d col %d data %s\n", target_row, c+1, temp_id);
                    }
                    tally_count++;
                }
                ptr += len;
            }
            state = 1; // 切換狀態：準備尋找下方對應粒子數的數據
            continue;
        }
        
        // 2. 抓取對應粒子數 (NPS) 的數據列
        long long nps_check;
        if (sscanf(line, "%lld", &nps_check) == 1 && nps_check == particles && state == 1) {
            int data_group_idx = 0; 
            char *token = strtok(line, " \t\n");
            token = strtok(NULL, " \t\n"); // 跳過 NPS 欄位本身

            // 依序填入各個 Tally 組的數據
            while (token != NULL && data_group_idx < groups) {
                int target_row = count * groups + data_group_idx + 1;
                if (target_row >= dt->rows) {
				    fclose(fp);
				    return -1;
				}
                for (int c = 0; c < set; c++) {
                    if (token != NULL) {
                        dt->table[target_row][c+1].type = TYPE_STRING;
                        dt->table[target_row][c+1].data.s = strdup(token);
                        // printf("get row %d col %d data %s\n", target_row, c+1, token);
                        token = strtok(NULL, " \t\n");
                    }
                }
                data_group_idx++;
            }

            count++;   // 完成一組 (包含所有 groups) 的處理
            state = 0; // 回到尋找下一個 Tally 區塊的狀態
            continue;
        }
    }
    
    fclose(fp);
    return 0;
}

/**
 * @brief 產生最終彙總報表 CSV
 * @details 將 source.csv 的基礎資訊與多個 .o 檔解析出的數據橫向合併。
 * @param folder  輸出資料夾名稱
 * @param origin  來自 source.csv 的原始表格
 * @param myData  包含所有檔案數據的 TableSet
 * @param myFiles 檔名清單 (用於提取標頭)
 */
int exportFinalReport(const char* folder, DynamicTable* origin, TableSet* myData, char** myFiles) {
    char out_path[512];
    snprintf(out_path, sizeof(out_path), "%s/Final_Report.csv", folder);
    FILE *fp = fopen(out_path, "w");
    if (!fp) return -1;

    fputs("\xEF\xBB\xBF", fp); 

    // --- Row 1: 標頭 ---
    fprintf(fp, "%s,,%s\n", 
            (origin->rows > 0 && origin->table[0][0].data.s) ? origin->table[0][0].data.s : "", 
            folder);

    // --- Row 2: 提取檔名數值 ---
    fprintf(fp, ",,"); 
    for (int i = 0; i < myData->file_count; i++) {
        char temp_name[256];
        if (myFiles[i]) {
            strncpy(temp_name, myFiles[i], sizeof(temp_name)-1);
            temp_name[sizeof(temp_name)-1] = '\0';
            char *dot = strrchr(temp_name, '.');
            if (dot) *dot = '\0';
            char *last_dash = strrchr(temp_name, '-');
            char *pure_val = (last_dash) ? last_dash + 1 : temp_name;
            fprintf(fp, "%s", pure_val);
        }
        // 根據欄位數補逗號 (dt->cols + 1 是為了檔案間的間隔)
        int col_count = (myData->tables[i].rows > 0) ? myData->tables[i].cols : 1;
        for (int c = 0; c <= col_count; c++) fprintf(fp, ",");
    }
    fprintf(fp, "\n");

    // --- Row 3: 輸出細項標籤 (no., mean, error...) ---
    fprintf(fp, ",,");
    for (int f = 0; f < myData->file_count; f++) {
        DynamicTable *dt = &(myData->tables[f]);
        if (dt->rows > 0 && dt->table != NULL) {
            for (int c = 0; c < dt->cols; c++) {
                fprintf(fp, "%s,", (dt->table[0][c].data.s) ? dt->table[0][c].data.s : "");
            }
        } else {
            fprintf(fp, "NO_DATA,");
        }
        fprintf(fp, ","); 
    }
    fprintf(fp, "\n");

    // --- Row 4 之後: 對齊數據列 (解決長度不一問題) ---
    
    // 1. 計算全局最大行數
    int max_rows = origin->rows;
    for (int f = 0; f < myData->file_count; f++) {
        if (myData->tables[f].rows > max_rows) {
            max_rows = myData->tables[f].rows;
        }
    }

    // 2. 以 max_rows 為基準輸出，r 從 1 開始 (跳過各表的標頭)
    for (int r = 1; r < max_rows; r++) {
        
        // --- A. 輸出左側 origin 數據 (來源 CSV) ---
        if (r < origin->rows && origin->table[r] != NULL) {
            const char *cell0 = (origin->table[r][0].type == TYPE_STRING) ? origin->table[r][0].data.s : "";
            const char *cell1 = (origin->table[r][1].type == TYPE_STRING) ? origin->table[r][1].data.s : "";
            fprintf(fp, "%s,%s", cell0 ? cell0 : "", cell1 ? cell1 : "");
        } else {
            fprintf(fp, ",");
        }

        // --- B. 橫向輸出各個解析檔案的數據 ---
        for (int f = 0; f < myData->file_count; f++) {
            DynamicTable *dt = &(myData->tables[f]);
            
            if (r < dt->rows && dt->table[r] != NULL) {
                // 該檔案在此行有數據
                for (int c = 0; c < dt->cols; c++) {
                    fprintf(fp, ",");
                    if (dt->table[r][c].type == TYPE_STRING) {
                        fprintf(fp, "%s", dt->table[r][c].data.s ? dt->table[r][c].data.s : "");
                    } else {
                        fprintf(fp, "%g", dt->table[r][c].data.d); 
                    }
                }
                fprintf(fp, ","); // 檔案間隔空欄
            } else {
                // 該檔案較短或無數據，補齊該表應佔的欄位
                int col_to_fill = (dt->rows > 0) ? dt->cols : 1;
                for (int c = 0; c <= col_to_fill; c++) fprintf(fp, ",");
            }
        }
        fprintf(fp, "\n");
    }

    fclose(fp);
    printf(" >> 報告產出完成，檔案最多有 %d 筆數據\n", max_rows);
    return 0;
}
