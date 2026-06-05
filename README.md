# Daily Tracker — 每日追蹤主機項目

每日自動從 EPB 查詢各門市保固搭售率，並掃描 Apple Mail 中的 Personal Setup 信件，產出含 Mysetup 提交率的 Excel 報表。

## 功能

- **EPB 查詢**：Mac / iPhone / iPad / Watch / AirPods 主機台數、SACare、AppleCare+
- **Mysetup 提交率**：自動掃描 Apple Mail 中每日寄出的 *Personal Setup: Setup Data* 信件 Excel 附件，統計各門市提交筆數
- **雙表輸出**：週圍區間 + 月累積，提交率 < 60% 紅字標示
- **支援兩個區域**：北一區（士林/微風/美麗華/阿波羅/高島屋/羅東）、北二區（永和/板橋誠品/西門/花蓮/板橋遠百/新莊宏匯/新店裕隆城）

## 環境需求

- macOS（需要 Apple Mail 本機儲存信件）
- Python 3.9+
- JDK 1.8（EPB 連線用）
- EPB App 資料夾放在 `~/Desktop/北一區週報-app/`

## 安裝

```bash
pip install -r requirements.txt
```

## 設定

開啟 `config.py`，依照自己的環境修改：

```python
# Java 路徑（確認版本）
JAVA = "/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home/bin/java"

# EPB App 資料夾（預設放在 ~/Desktop/）
EPB_APP_DIR = Path("~/Desktop/北一區週報-app").expanduser()

# 報表輸出目錄（預設 ~/daily-tracker-output/）
OUTPUT_BASE = Path("~/daily-tracker-output").expanduser()
```

## 執行

```bash
# 北一區 — 自動（最近完整週）
python3 warranty_report_n1.py

# 北一區 — 指定日期區間
python3 warranty_report_n1.py 5/31 6/4

# 北二區
python3 warranty_report_n2.py 5/31 6/4
```

輸出檔案位置：
- `~/daily-tracker-output/北一區/保固搭售率_MMDD-MMDD.xlsx`
- `~/daily-tracker-output/北二區/保固搭售率北二區_MMDD-MMDD.xlsx`

## Mysetup 信件說明

每天早上 10 點會收到主旨為 `Personal Setup: Setup Data DD-M月-YY - DD-M月-YY` 的信件，附件為 Excel，腳本會自動解析並計算各門市提交率。

北一區 Mysetup 統計門市：士林、微風、美麗華、阿波羅、高島屋（羅東不列入）  
北二區 Mysetup 統計門市：西門、板橋遠百、新莊宏匯、新店裕隆城
