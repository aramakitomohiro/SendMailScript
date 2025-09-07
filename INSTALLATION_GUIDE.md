# 🚀 インストール・セットアップガイド

## 📋 **前提条件**

### **システム要件**
- **OS**: Windows Server 2016以降 / Windows 10以降
- **PowerShell**: バージョン5.1以降
- **メール環境**: Exchange Server または SMTP サーバー
- **ネットワーク**: ファイルサーバー (`goa-mailsv`) へのアクセス権限

### **必要な権限**
- ローカル管理者権限
- ネットワークドライブへの読み書き権限
- SMTPサーバーへの送信権限

## 📁 **ディレクトリ構成設定**

### **1. 基本ディレクトリ作成**

```powershell
# 管理者権限でPowerShellを起動し、以下のコマンドを実行

# ローカルディレクトリ作成
New-Item -ItemType Directory -Path "C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails" -Force
New-Item -ItemType Directory -Path "C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\log" -Force
New-Item -ItemType Directory -Path "C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\BK" -Force

# ネットワークドライブ作成（事前にファイルサーバー管理者に依頼）
# \\goa-mailsv\UserData
# \\goa-mailsv\001_InspectingAndMovingInputData
# \\goa-mailsv\001_InspectingAndMovingInputData\log
```

### **2. スクリプトファイル配置**

```
📁 C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\
├── 📄 BulkSendingExternalEmail.ps1        # メイン送信スクリプト
├── 📄 BulkSendingExternalEmail.psm1       # 共通関数モジュール
└── 📄 BulkSendingExternalEmail-Test.ps1   # テストスクリプト

📁 \\goa-mailsv\001_InspectingAndMovingInputData\
├── 📄 InspectingAndMovingInputData.ps1    # CSV前処理スクリプト
└── 📄 mailsettings_本番用.ini             # 設定ファイル
```

## ⚙️ **設定ファイル構成**

### **mailsettings_本番用.ini の編集**

```ini
[General]
# 送信前確認ダイアログ
Confirm=true

# 送信元メールアドレス（実際のアドレスに変更）
From=your-system@company.co.jp

# BCCアドレス（監査用、実際のアドレスに変更）
Bcc=audit@company.co.jp

[Paths]
# ユーザー領域（CSV配置場所）
UserData=\\goa-mailsv\UserData

# 準備領域
001_InspectingAndMovingInputData=\\goa-mailsv\001_InspectingAndMovingInputData

# 送信処理領域
002_BulkSendingExternalEmails=C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails

# ログ格納先
001_InspectingAndMovingInputDataLog=\\goa-mailsv\001_InspectingAndMovingInputData\log
002_BulkSendingExternalEmailsLog=C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\log

# バックアップ領域
CSV_FileBackup=C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\BK

[SMTP]
# SMTPサーバー設定（実際のサーバー名に変更）
Server=your-exchange-server
Port=25

# ファイル種別定義
[KOZA_KAISETSU]
CsvKeyword=口座_メール送信時に使用

[GAKKO_JIFU]
CsvKeyword=学校_メール送信時に使用

[TODOKE_HENKO]
CsvKeyword=届出変更_メール送信時に使用
```

## 🔧 **JP1設定**

### **バッチジョブ定義**

```bash
# JP1/AJS バッチジョブ定義例
#!/bin/bash
# ジョブ名: BulkEmailSending
# 実行間隔: 5分
# 実行コマンド:
PowerShell.exe -ExecutionPolicy Bypass -File "C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\BulkSendingExternalEmail.ps1"

# 戻り値判定:
# 0: 正常終了（成功）
# 1: 処理対象なし（正常）
# 2: 一部失敗（警告）
# 3: 致命的エラー（異常）
```

### **監視設定**

```xml
<!-- JP1/IM 監視定義例 -->
<MonitorRule>
    <Name>BulkEmailSending_Monitor</Name>
    <Description>メール送信処理監視</Description>
    <Condition>
        <ExitCode>3</ExitCode>
        <LogPattern>ERROR</LogPattern>
    </Condition>
    <Action>
        <SendAlert>true</SendAlert>
        <NotifyUsers>admin@company.co.jp</NotifyUsers>
    </Action>
</MonitorRule>
```

## 🧪 **動作テスト手順**

### **1. 基本動作テスト**

```powershell
# テストデータ作成
cd "C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails"
.\BulkSendingExternalEmail-Test.ps1 -CreateTestData

# 設定ファイルテスト
.\BulkSendingExternalEmail-Test.ps1 -RunFullTest

# テストデータクリーンアップ
.\BulkSendingExternalEmail-Test.ps1 -CleanupOnly
```

### **2. CSV形式テスト**

```csv
"メール送信日","メールアドレス","メール件名","メール本文"
"","test@example.com","テスト件名","テスト本文です。"
```

**ファイル名**: `口座_メール送信時に使用_2024_03_15 14_30_00.csv`
**エンコーディング**: UTF-8 BOM付き

### **3. エラーパターンテスト**

```powershell
# 各種エラーパターンを順次テスト
# 1. ファイル名形式エラー
# 2. エンコーディングエラー
# 3. 必須カラム不足
# 4. メールアドレス形式エラー
```

## 🔐 **セキュリティ設定**

### **PowerShell実行ポリシー**

```powershell
# 管理者権限で実行
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

# 確認
Get-ExecutionPolicy -List
```

### **ファイル権限設定**

```
📁 ローカルディレクトリ
├── SYSTEM: フルコントロール
├── Administrators: フルコントロール
└── JP1実行ユーザー: 読み取り・実行

📁 ネットワークドライブ
├── Domain Admins: フルコントロール
├── メール送信システム用グループ: 変更
└── 一般ユーザー: 読み取り専用（UserDataのみ書き込み可）
```

## 📊 **ログ設定**

### **ログファイルパス**
- **スクリプトA**: `\\goa-mailsv\001_InspectingAndMovingInputData\log\A_yyyy_MM_dd.log`
- **スクリプトB**: `C:\005_bulkSendingExternalEmails\002_BulkSendingExternalEmails\log\B_yyyy_MM_dd.log`

### **ログローテーション**
- **保存期間**: 731日（2年間）
- **自動削除**: 各スクリプト起動時に実行
- **ファイル形式**: UTF-8

## ⚡ **パフォーマンス最適化**

### **推奨設定**

```powershell
# .NET Framework ガベージコレクション最適化
$env:COMPlus_gcServer = 1
$env:COMPlus_gcConcurrent = 1

# PowerShell プロセス優先度設定
$process = Get-Process -Id $PID
$process.PriorityClass = 'High'
```

### **システムリソース監視**

```powershell
# リソース使用量確認スクリプト
Get-Counter -Counter "\Memory\Available MBytes"
Get-Counter -Counter "\Processor(_Total)\% Processor Time"
Get-WmiObject -Class Win32_LogicalDisk | Select-Object DeviceID, FreeSpace, Size
```

## 🚨 **トラブルシューティング**

### **よくある問題と解決策**

| 問題 | 原因 | 解決策 |
|------|------|--------|
| CSV読み込みエラー | エンコーディング不正 | UTF-8 BOM付きで保存し直し |
| SMTP接続エラー | ネットワーク設定 | サーバー名・ポート番号確認 |
| ファイル権限エラー | アクセス権限不足 | 管理者に権限設定依頼 |
| メモリ不足 | 大容量CSV処理 | CSVファイル分割 |

### **ログ確認コマンド**

```powershell
# エラーログ抽出
Get-Content "C:\...\log\B_2024_03_15.log" | Where-Object { $_ -match "ERROR" }

# 送信件数集計
Get-Content "C:\...\log\B_2024_03_15.log" | Where-Object { $_ -match "送信成功" } | Measure-Object
```

## 📞 **サポート情報**

### **連絡先**
- **システム管理者**: admin@company.co.jp
- **技術サポート**: support@company.co.jp
- **緊急時**: emergency@company.co.jp

### **保守スケジュール**
- **定期メンテナンス**: 毎月第2土曜日 2:00-4:00
- **緊急メンテナンス**: 24時間前通知
- **システム更新**: 四半期ごと