# salesforce-xlsx-to-customfiled

Upsert CustomField and FieldPermissions with Excel.

This is [Electron](https://www.electronjs.org/)-based application.

Salesforce 用のエクセルファイルからカスタム項目を作るツールです。

もともと Node.js でコマンドラインを使って動かしていた社内向けに作ったツールです。

# Platforms

Windows 64bit.

# Usage

## Quick Start

1. [Download Zip](./)
2. Execute "salesforce-upsert-customfiled-tool.exe".
3. Input login information.
4. Choose "Samples/MinimumSample_Account.xlsx".
5. Press "Execute" button.
6. Wait all permission result.

## Define Customfield in .xlsx

There are sample definitions in "samples/CustomFieldTest\_\_c.xlsx" file.

samples フォルダのエクセルに書き方のサンプルがあります。

## Retry Execute

If you defined formula type field, perhaps the field will be error.

項目の UPSERT は 10 件ずつ一括処理をするため数式などはエラーになりやすいです。そのときは再実行すると通ります。

# Usage Command Line

```
node sfuce.js sample.xlsx
```

# License

[MIT](/LICENSE)

Copyright (c) 2021 Takahiro Komori
