# Daily Tracker — 每日追蹤主機項目（北二區）

每日自動從 EPB 查詢各門市保固搭售率，並掃描 Apple Mail 中的 Personal Setup 信件，產出含 Mysetup 提交率的 Excel 報表。

## 快速開始

### 1. 下載
```bash
git clone https://github.com/Harveyzxc15/daily-tracker
cd daily-tracker
```

### 2. 一鍵安裝
在 Finder 中**雙擊 `setup.command`**，腳本會自動：
- 安裝 Python 套件
- 確認 Java 與 EPB 環境
- 建立輸出目錄

### 3. 執行
```bash
python3 warranty_report_n2.py              # 自動（最近完整週）
python3 warranty_report_n2.py 5/31 6/4    # 指定日期區間
```

輸出檔案：`~/daily-tracker-output/北二區/保固搭售率北二區_MMDD-MMDD.xlsx`

---

## 環境需求

| 項目 | 說明 |
|------|------|
| macOS | 需要 Apple Mail 本機儲存信件 |
| Python 3.9+ | `pip3 install -r requirements.txt` |
| JDK 1.8 | EPB 連線用，路徑：`/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/` |
| EPB App | 放在 `~/Desktop/北一區週報-app/` |

## 設定（如需調整路徑）

開啟 `config.py`：

```python
JAVA        = "/Library/Java/.../bin/java"      # Java 路徑
EPB_APP_DIR = Path("~/Desktop/北一區週報-app")   # EPB App 位置
OUTPUT_BASE = Path("~/daily-tracker-output")     # 輸出根目錄
```

## 門市說明

**北二區門市（`warranty_report_n2.py`）**

| 門市 | Mysetup 統計 |
|------|:------------:|
| 永和 | — |
| 板橋誠品 | — |
| 西門 | ✅ |
| 花蓮 | — |
| 板橋遠百 | ✅ |
| 新莊宏匯 | ✅ |
| 新店裕隆城 | ✅ |

> Mysetup 提交率 < 60% 顯示紅字

## 報表內容

每份 Excel 包含兩張表：
- **週圍**：指定起迄日的保固搭售率 + Mysetup 提交率
- **月累積**：當月 1 日至指定結束日的累積數字

欄位：Mac / iPhone / iPad / Watch / AirPods 台數、ACPP+、SACare、搭售率、ARpedia 數量、喇叭金額、Mysetup 提交數、Mysetup 提交率
