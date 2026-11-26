# RIS-CombineReport

CT 多筆檢查自動合併報告與影像的小工具。

這個工具會操作 RIS（放射線資訊管理系統）「登錄作業清單」的畫面。同一位病人有多筆 CT 時，它會自動在清單上幫忙合併報告和影像。

## 檔案

| 檔案 | 用途 |
|---|---|
| `ct_auto.py` | **CT 檢查室版**：只有自動模式，畫面上只能改「刷新時間」 |
| `ct_merge.py` | 完整版，有手動模式和自動模式。`ct_auto.py` 也要靠它 |
| `ct_merge_no_merge.txt` | 不用合併的醫令清單 |
| `build_portable.ps1` | 做免安裝版（不用另外安裝 Python） |
| `portable_make_shortcut.py` | 給免安裝版在桌面建立捷徑 |

## 它會做什麼

同一位病人在清單上有 2 筆以上 CT 時：

- **合併序號不一樣**：用 Ctrl 選好那幾筆，按右鍵「合併報告」。
- **影像單號不一樣**：用 Ctrl 選好那幾筆，按右鍵「影像合併」。影像還沒傳到的話，每隔「刷新時間」再試一次。
- **已經一樣**：不用做。

它不點開病人，也不按「開始檢查」或「檢查完成」，所以檢查室不會被換掉。

## 安全規則

RIS 是正式系統，按下去就直接生效，所以這個工具很保守：

- 只處理「全部都是 CT」的病人。有非 CT 檢查（X 光等）的，整位病人都不碰。
- 醫令符合 `ct_merge_no_merge.txt` 的病人整位不碰。
- 報告醫師不一樣的病人不碰，留給人工處理。
- 鍵盤和滑鼠 20 秒沒人動才開始動作，有人在用 RIS 就先等。
- RIS 跳出沒看過的視窗時，程式會停下來，不會自己亂按。
- 記錄寫在 `ct_merge_log.txt`，只記照會單號、醫令和醫師，不記病人姓名和病歷號。

## 使用方式

### 免安裝版（放到 CT 檢查室）

1. 在有安裝 Python 的電腦執行：
   ```
   powershell -ExecutionPolicy Bypass -File build_portable.ps1
   ```
   會產生 `portable_build\CT自動合併_免安裝版.zip`。
2. 把 zip 拿到 CT 檢查室的電腦解壓縮，雙擊「建立桌面捷徑.bat」一次。
3. 在 RIS 開「登錄作業清單」，選好這間檢查室。
4. 用桌面上的「CT自動合併」開啟程式，按「開始」。

以後要更新程式，只要換掉 `ct_auto.py` 和 `ct_merge.py` 這兩個檔案。

### 直接用 Python 執行

需要 Windows 和 Python 3.12，先安裝套件：

```
pip install pywinauto==0.6.9 comtypes==1.4.16 pywin32==312 six==1.17.0
```

接著用下面其中一個指令開啟：

```
python ct_auto.py            # CT 檢查室版
python ct_merge.py           # 完整版，開啟時選手動或自動
python ct_merge.py auto      # 完整版，直接用自動模式
python ct_merge.py manual    # 完整版，直接用手動模式
```

## 不用合併的醫令

`ct_merge_no_merge.txt` 一行寫一個醫令名稱，大小寫和多餘的空白都不管：

- `#` 開頭的是說明。
- `*` 代表任意文字，例如 `*(HMC)` 代表所有 (HMC) 結尾的醫令。預設就是這一行。
- `!` 開頭是例外，代表這個醫令要合併，例如 `!CT-Chest(HMC)`。

病人的 CT 只要有一筆符合，這位病人就整位不碰。
