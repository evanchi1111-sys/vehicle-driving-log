# vendor（第三方程式檔）

網頁使用的兩個第三方工具，直接放在專案內，`index.html` 不需要從外部網站載入。

| 檔案 | 用途 | 版本 | 原始來源 | 授權 |
|---|---|---|---|---|
| `qrcode.min.js` | 產生 QR Code | qrcodejs 1.0.0 | https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js | MIT（作者 davidshimjs，https://github.com/davidshimjs/qrcodejs） |
| `xlsx.full.min.js` | 匯入單位與使用人名單（Excel） | SheetJS（xlsx）0.18.5 | https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js | Apache-2.0（SheetJS，https://sheetjs.com） |

SHA-256（用來確認檔案沒有被改動）：

```
c541ef06327885a8415bca8df6071e14189b4855336def4f36db54bde8484f36  qrcode.min.js
c9506197caf809a075b6dee1da0d36fb19da7158ffe8a88e7b0c96c5d8623c99  xlsx.full.min.js
```

- 這兩個檔案是原樣放入，沒有修改內容。
- `app.template.html`（Claude Artifact 版）仍然使用上表的來源網址，因為 Artifact 沒有 `vendor/` 資料夾。
- 要更新版本時：換掉這裡的檔案、更新上表的版本與 SHA-256，再執行 `build.ps1` 重新產生 `index.html`。
