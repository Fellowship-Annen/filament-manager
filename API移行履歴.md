# API移行履歴

## 2026-10-09

### 新API

`https://script.google.com/macros/s/AKfycby24DTSic0qx2Y2ZqNU2RsaEuXPRSQRwdBPgshbUPI2ui8vPY1PUviEKD9QeMx5gIe3Bg/exec`

- コピー後の新スプレッドシートに紐付けて作成。
- `filament-form-history-v3` のhealth応答を確認済み。
- 重量更新、場所変更、入庫登録、印刷回数、ラベル注釈、フィラメントマスター取得、メーカーマスター取得、フィラメントマスター追加に対応。

### 旧API（切戻し用記録）

`https://script.google.com/macros/s/AKfycbzCCejDFlAfJV5sjxhzH7wp1-4S7oVK6s0Gru_m4QXVmSRdnnOXDcCfqeogW5GA4PFR/exec`

### 公開前の確認事項

- 窓口のCSV取得先は旧スプレッドシートIDのままなので、新しいスプレッドシートのURLと在庫台帳のgidへ変更する。
- APIとCSVが同じスプレッドシートを参照していることを確認してからGitHubへ公開する。
