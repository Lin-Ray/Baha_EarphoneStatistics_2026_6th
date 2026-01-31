# BaHa Earphone Statistics

## 簡短說明
--
本專案目的在學習爬蟲處理巴哈姆特之`2026年初有線耳機普查兼前端普查——第六屆`相關的資料，從原始資料讀取，到輸出可供分析的 Excel 報表。主要腳本為 `BaHa_EarphoneStatistics.py`，產生範例輸出位於 `output/` 目錄。

## 功能項目
--
- 爬取使用目標網站(使用BeautifulSoup、requests)
- 匯出彙整報表為 Excel

## 安裝與執行
--
建議使用 Python 3.8+，專案相依請參考 `requirements.txt`。
```
# 安裝相依套件
pip install -r requirements.txt

# 執行主程式
python BaHa_EarphoneStatistics.py
```

## 輸出
- Excel 報表，預設儲存在 `output/BaHa_EarphoneStatistics_YYYYMMDD.xlsx`，包含統計摘要與原始資料彙整工作表

## 實作與細節

- 加入 logging 以記錄處理流程與錯誤（便於除錯）
- 爬蟲錯誤簡易對應 (替換UA[User Agents]、延遲請求、請求次數等)
- 網站動態物件取得 (例如:本範例要取得目標回文之`分享連結`，需要透過滑鼠移至Bar Menu才會顯示該按鈕)
- 分析內文格式並將分類標題做資料清洗 (例如:本範例該串主要定義`耳罩`、`耳塞`、`前端`，因資料源有部分未依據格式[例如: '耳罩'在本串有其他稱呼{耳罩式耳機, 耳罩式, 大耳'...等}]，所以必須將其做資料清洗統一以便輸出)


## 參考資料
--
[巴哈姆特-2026年初有線耳機普查兼前端普查——第六屆](https://forum.gamer.com.tw/C.php?page=1&bsn=60535&snA=28366)