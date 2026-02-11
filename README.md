# MRCP Data Toolkit v2.6

**MRCP Data Toolkit** 是一款專為核能與醫學物理工程研究設計的高效能數據整合工具。它能將 MRCP 模擬產生的海量文字輸出檔案（`.o` 檔）轉化為結構化的科研報表。透過 C 語言的強大解析能力與 PowerShell 的 Excel 自動化技術，大幅縮短數據整理時間。

## 📸 成果展示 (Result Preview)

![Project Result](result1.png)
> **圖 1：v2.5的GUI 批次選取介面與後台狀態監控** 。
![Project Result](result2.png)
> **圖 2：PowerShell 自動化生成的美化版 Excel 報表**。

---

## 🛠 運作流程 (Workflow)

1. 
**掃描 (Scanning)**：自動遍歷目錄，識別並按字母順序排列所有 `.o` 輸出檔案 。


2. **解析 (Parsing)**：使用狀態機模型從文字檔中精確提取粒子數 (`NPS`) 與各組 `Tally` 數據。
3. **校準 (Calibration)**：對比 `source.csv`，自動進行數據對齊、標籤清洗並剔除無效空值。
4. **整合 (Integration)**：驅動 Excel COM 物件，將多個資料夾數據彙整至獨立分頁並執行樣式美化。

---

## 🚀 核心功能與技術特色

* **GUI 批次處理 (v2.6 新特性)**：
* 基於 Win32 API 打造輕量化視窗 。


* 支援 `Ctrl/Shift` **多選資料夾**，一次處理多組實驗數據 。


* 彩色控制台日誌輸出，即時監控處理進度與系統狀態 。




* **高效解析引擎**：
* **狀態機解析**：穩定處理大型輸出檔，精確定位 `Tally` 數據區塊。
* **智能校準機制**：自動偵測數據間的空格長度，若數據缺失則自動剔除無效標籤並往前遞補，確保報表嚴整。


* **自動化報表美化 (PowerShell 驅動)**：
* **多工作表彙整**：每個資料夾自動歸納為 Excel 中的一個獨立分頁 (Sheet)。
* **科研格式標準**：全自動設定 **Times New Roman** 字體、自動調整欄寬。
* **視覺化標註**：自動偵測標頭區域並填充**黃色底色**，提升閱讀效率。


* **健壯性設計**：
* 採用 **動態記憶體管理 (Union & Dynamic Table)**，優化大規模數據集的內存佔用。
* 內建 **UTF-8 BOM** 處理邏輯，徹底解決 Excel 開啟 CSV 時的亂碼困擾。



---

📁 檔案結構說明 

| 檔案名稱 | 說明 |
| --- | --- |
| **`main.exe`** | 核心程式：包含 Win32 GUI 介面與批次調度邏輯。 |
| **`tools.h / config.h`** | 演算法庫：負責資料結構定義、狀態機解析與內存管理。 |
| **`merge.ps1`** | 自動化腳本：負責調用 Excel COM 執行多表整合與排版。 |
| **`source.csv`** | 基準參數表：定義數據對齊的原始欄位與基礎資訊。 |

---

## 🛠️ 安裝與開發環境

* **作業系統**：Windows 10 / 11
* **軟體依賴**：必須安裝 **Microsoft Excel** (用於執行 `merge.ps1` 自動化功能)。
* **編譯環境**：
* 編譯器：Dev-C++ / MinGW-w64 / GCC 。


* 連結標記：`-lcomdlg32 -lshell32 -lgdi32` 。


---

## 📖 使用手冊 (Quick Start)

### 1. 環境佈置

將 `main.exe`、`merge.ps1` 與 `source.csv` 置於同一目錄，並將模擬數據（`.o` 檔）按資料夾分類放置於子目錄中 。

### 2. 啟動與解析

執行 `main.exe`。若未偵測到 `source.csv`，程式會自動提示手動選取 。在介面清單中選取目標資料夾後，點擊 **「批次處理」** 。

### 3. 一鍵生成報表

解析完成後，系統會詢問「是否執行整合美化」。點擊**「是」**，系統將調用 PowerShell 在背景完成所有 Excel 排版工作，最終生成 `MRCP_Total_Summary.xlsx` 。

---

## 📜 授權條款 (License)

本專案採用 **MIT License** 授權。

**KikKoh @2026** - *Dedicated to Nuclear Engineering Data Solutions.*

---

### 💡 開發者建議

1. **路徑檢查**：確保執行路徑中不包含特殊字元，以維持 PowerShell 腳本的穩定性。
2. **Excel 進程**：若腳本執行中斷，請檢查工作管理員中是否有殘留的 Excel 進程。

Would you like me to refine the technical documentation for any of the specific C functions, such as the `loadSingleFileData` state machine logic?
