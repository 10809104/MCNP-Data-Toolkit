# MCNP Data Toolkit v2.5

**MCNP Data Toolkit** 是一款專為核能工程研究設計的高效能數據整合工具。它能將 MCNP 模擬產生的海量文字輸出檔案（`.o` 檔）轉化為結構化的科研報表。透過 C 語言的強大解析能力與 PowerShell 的 Excel 自動化技術，大幅縮短數據整理時間。

## 📸 成果展示 (Result Preview)

![Project Result](result1.png)
> **圖：自動化處理展示**。
![Project Result](result2.png)
> **圖：自動化處理結果**
---

## 🛠 運作流程 (Workflow)

1. **掃描 (Scanning)**：遍歷資料夾，識別並排序所有 MCNP 輸出檔案。
2. **解析 (Parsing)**：使用狀態機模型從文字檔中提取 `NPS` 與各類 `Tally` 數據。
3. **校準 (Calibration)**：對比 `source.csv`，自動進行數據列對齊與標籤清洗。
4. **整合 (Integration)**：PowerShell 驅動 Excel COM 物件，執行分頁彙整與樣式美化。

---

## 🚀 核心功能與技術特色

* **GUI 批次處理**：基於 Win32 API 打造輕量化視窗，支援 `Ctrl/Shift` 多選資料夾，解決傳統手動指令的低效問題。
* **高效解析引擎**：
* **狀態機解析**：穩定處理數百 MB 的輸出檔，精確定位數據區塊。
* **智能校準**：自動偵測並過濾數據中的空白行、無效標籤，確保最終 CSV 格式嚴整。


* **自動化報表美化**：
* **多工作表彙整**：每個處理的資料夾自動成為 Excel 中的一個獨立分頁 (Sheet)。
* **科研格式標準**：全自動設定 **Times New Roman** 字體、自動調整欄寬。
* **視覺化標註**：自動為數據標頭（Row 2）填充顏色，提升閱讀效率。


* **健壯性設計**：
* 採用 **動態記憶體管理**，能處理數萬行規模的大型數據集。
* 內建 **UTF-8 BOM** 處理邏輯，徹底解決 Excel 開啟 CSV 時的中文亂碼困擾。



---

## 📁 檔案結構說明

| 檔案名稱 | 說明 |
| --- | --- |
| **`main.exe`** | 核心編譯程式，包含 GUI 介面與批次調度邏輯。 |
| **`tool.h / config.h`** | 核心演算法庫，負責資料結構定義與狀態機解析。 |
| **`merge.ps1`** | 自動化腳本，負責與 Microsoft Excel 溝通執行美化。 |
| **`source.csv`** | 基準參數表，定義了數據對齊的原始欄位。 |

---

## 🛠️ 安裝與開發環境

* **作業系統**：Windows 10 / 11
* **軟體依賴**：必須安裝 **Microsoft Excel** (用於 Excel 自動化功能)。
* **編譯環境**：
* 編譯器：Dev-C++ / MinGW-w64 / GCC
* 編譯標記：`-lcomdlg32 -lshell32 -lgdi32`



---

## 📖 使用手冊 (Quick Start)

### 1. 環境佈置

將 `main.exe`、`merge.ps1` 與 `source.csv` 置於同一目錄，並將模擬數據（`.o` 檔）按資料夾分類放置。

### 2. 啟動與解析

執行 `main.exe`。若根目錄無 `source.csv`，程式會自動彈出對話框請你指定來源。隨後在列表框選取資料夾，點擊 **「批次處理」**。

### 3. 一鍵生成報表

解析完成後，系統會詢問「是否執行整合美化」。點擊**「是」**，Excel 會在背景啟動並自動完成所有排版工作，最終生成 `MRCP_Total_Summary.xlsx`。

---

## 📜 授權條款 (License)

本專案採用 **MIT License** 授權。

**KikKoh @2026** - *Dedicated to Nuclear Engineering Data Solutions.*

---

### 💡 開發者建議

1. **Release 頁面**： `main.exe` 已壓成 `.zip` 並附上範例。
