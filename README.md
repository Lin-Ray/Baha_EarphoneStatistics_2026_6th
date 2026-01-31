## BaHa Earphone Statistics

簡短說明
--
本專案用於自動化處理耳機相關的測試或使用資料，從原始資料讀取、清理、計算統計指標，到輸出可供分析的 Excel 報表。主要腳本為 `BaHa_EarphoneStatistics.py`，產生範例輸出位於 `output/` 目錄。

主要功能
--
- 讀取原始資料（支援 Excel/CSV，視程式實作而定）
- 資料清理與欄位正規化（日期轉換、缺值處理）
- 計算統計指標（平均、標準差、分位數、事件計數等）
- 匯出彙整報表為 Excel（多工作表、摘要與明細）
- 支援基本參數化（輸入檔案、輸出目錄、日期篩選等）

安裝與執行（範例）
--
建議使用 Python 3.8+，專案相依請參考 `requirements.txt`。

PowerShell 範例：
```powershell
# 建立並啟用虛擬環境（PowerShell）
python -m venv .venv; .\.venv\Scripts\Activate.ps1

# 安裝相依套件
pip install -r requirements.txt

# 執行主程式（視程式參數而定）
python BaHa_EarphoneStatistics.py --input data/raw_data.xlsx --output output/ --start-date 2026-01-01 --end-date 2026-01-31
```

參數範例（請以實際程式實作為準）
--
- `--input` : 原始資料檔案路徑（CSV 或 XLSX）
- `--output` : 輸出目錄（預設為 `output/`）
- `--start-date` / `--end-date` : 篩選資料的日期區間

輸入 / 輸出
--
- 輸入：原始測試或使用資料（需包含時間戳、測量欄位等；請參考程式中所要求的欄位）
- 輸出：Excel 報表，預設儲存在 `output/BaHa_EarphoneStatistics_YYYYMMDD.xlsx`，包含統計摘要與原始資料彙整工作表

實作建議與細節
--
- 使用 pandas 處理資料讀取、清理與統計計算；使用 `pandas.ExcelWriter`（搭配 openpyxl）輸出多工作表
- 加入 logging 以記錄處理流程與錯誤（便於除錯）
- 在讀取前做欄位驗證，對必需欄位缺失給出明確錯誤訊息或紀錄到日誌

邊界情況與處理建議
--
- 空檔案或無有效列：輸出警告並跳過或生成空報表
- 欄位格式錯誤（例如日期解析失敗）：嘗試多種常用格式解析，仍失敗則記錄錯誤並略過該列或終止執行（視需求）
- 大型資料：採用分塊（chunk）讀取或增加記憶體限制，視執行環境調整

測試建議
--
- 建立最少兩組測資：正常資料（happy path）與包含缺值/格式錯誤的資料（edge cases）
- 驗證輸出檔案存在、工作表與欄位名稱正確、主要統計指標與預期一致

貢獻與授權
--
- 歡迎透過 Pull Request 提交改進；PR 請包含變更說明與相對應的測資或測試
- 專案授權請參考根目錄的 `LICENSE`

聯絡
--
如需協助或報告錯誤，請在 GitHub 專案頁面開 issue，並盡可能提供原始資料範例（去識別化）與錯誤輸出

最後備註
--
本 README 為一般性範本；若要更完整說明欄位名稱、參數細節或範例輸出，建議把欄位規格（schema）與範例資料一併放在 `docs/` 或 `examples/` 資料夾中。
