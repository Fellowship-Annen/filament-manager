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

- CSV取得先を、新スプレッドシートの在庫台帳（gid `93776902`）へ変更した。
- CSV 231行、必要列、新GASによる同一管理番号の取得一致を確認した。
- メーカーマスター17件、フィラメントマスター92件の取得を確認した。
- APIとCSVが同じスプレッドシートを参照していることを確認済み。

### 旧CSV（切戻し用記録）

`https://docs.google.com/spreadsheets/d/15UGqdqOwVQNx9qY02HtZIaqrD6Kd1dwmWuj4p3EGIGg/export?format=csv&gid=93776902`

### 新CSV

`https://docs.google.com/spreadsheets/d/1iG7fYgmrHHZy3cwKx9y9_4ddjI2qcGMeii3ekg8UoQc/export?format=csv&gid=93776902`
