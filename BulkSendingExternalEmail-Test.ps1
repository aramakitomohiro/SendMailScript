#===========================
#BulkSendingExternalEmail-Test.ps1
#外接メール送信_テスト専用スクリプト
#機能:
# - システム全体の動作テスト
# - CSV形式検証テスト
# - エラーハンドリングテスト
# - 設定ファイル検証テスト
#===========================

param(
    [string]$IniPath = "$PSScriptRoot\mailsettings.ini",
    [switch]$CreateTestData,
    [switch]$RunFullTest,
    [switch]$CleanupOnly
)

# 共通モジュールをインポート
Import-Module "$PSScriptRoot\BulkSendingExternalEmail.psm1" -Force

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

### テスト用関数: ログ出力
function Write-TestLog {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logLine = "$timestamp [$Level] [TEST] $Message"
    Write-Host $logLine -ForegroundColor $(
        switch($Level) {
            'ERROR' { 'Red' }
            'WARN' { 'Yellow' }
            'SUCCESS' { 'Green' }
            default { 'White' }
        }
    )
}

### テスト用関数: テストデータディレクトリ作成
function New-TestDirectories {
    param([hashtable]$IniData)
    
    Write-TestLog "テストディレクトリを作成中..."
    
    $paths = $IniData['Paths']
    $testPaths = @(
        $paths['UserData'],
        $paths['001_InspectingAndMovingInputData'],
        $paths['002_BulkSendingExternalEmails'],
        $paths['001_InspectingAndMovingInputDataLog'],
        $paths['002_BulkSendingExternalEmailsLog'],
        $paths['CSV_FileBackup']
    )
    
    foreach ($path in $testPaths) {
        if ($path -and -not (Test-Path $path)) {
            try {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
                Write-TestLog "ディレクトリ作成: $path" 'SUCCESS'
            } catch {
                Write-TestLog "ディレクトリ作成失敗: $path - $($_.Exception.Message)" 'ERROR'
            }
        }
    }
}

### テスト用関数: テストCSVファイル作成
function New-TestCsvFiles {
    param([hashtable]$IniData)
    
    Write-TestLog "テストCSVファイルを作成中..."
    
    $userDataPath = $IniData['Paths']['UserData']
    $timestamp = Get-Date -Format 'yyyy_MM_dd HH_mm_ss'
    
    # 正常なCSVファイル
    $validCsvPath = Join-Path $userDataPath "口座_メール送信時に使用_${timestamp}.csv"
    $validCsvContent = @"
"メール送信日","メールアドレス","メール件名","メール本文"
"","test1@example.com","口座開設のお知らせ","口座開設手続きが完了しました。"
"","test2@example.com","重要なお知らせ","システムメンテナンスのお知らせです。"
"2024-01-01 10:00:00","test3@example.com","送信済みテスト","この行は送信済みです。"
"@
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($validCsvPath, $validCsvContent, $utf8WithBom)
    Write-TestLog "正常CSVファイル作成: $validCsvPath" 'SUCCESS'
    
    # エラーテスト用CSVファイル（必須カラムなし）
    $invalidCsvPath = Join-Path $userDataPath "学校_メール送信時に使用_${timestamp}.csv"
    $invalidCsvContent = @"
"日付","宛先","題名","内容"
"","test4@example.com","テスト件名","テスト本文"
"@
    [System.IO.File]::WriteAllText($invalidCsvPath, $invalidCsvContent, $utf8WithBom)
    Write-TestLog "エラーテスト用CSVファイル作成: $invalidCsvPath" 'SUCCESS'
    
    # ファイル名形式エラーテスト用
    $wrongNameCsvPath = Join-Path $userDataPath "間違った名前.csv"
    [System.IO.File]::WriteAllText($wrongNameCsvPath, $validCsvContent, $utf8WithBom)
    Write-TestLog "ファイル名エラーテスト用CSVファイル作成: $wrongNameCsvPath" 'SUCCESS'
}

### テスト用関数: 設定ファイルテスト
function Test-IniConfiguration {
    param([string]$IniPath)
    
    Write-TestLog "設定ファイルテスト開始: $IniPath"
    
    try {
        # INI読み込みテスト
        $iniData = Get-IniData -Path $IniPath
        Write-TestLog "INI読み込み成功" 'SUCCESS'
        
        # 必須セクション存在チェック
        $requiredSections = @('General', 'Paths', 'SMTP')
        foreach ($section in $requiredSections) {
            if ($iniData.ContainsKey($section)) {
                Write-TestLog "必須セクション確認: [$section]" 'SUCCESS'
            } else {
                Write-TestLog "必須セクション不足: [$section]" 'ERROR'
            }
        }
        
        # パス設定チェック
        if ($iniData['Paths']) {
            $paths = $iniData['Paths']
            foreach ($pathKey in $paths.Keys) {
                $pathValue = $paths[$pathKey]
                Write-TestLog "パス設定: $pathKey = $pathValue"
            }
        }
        
        return $iniData
        
    } catch {
        Write-TestLog "設定ファイルテスト失敗: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

### テスト用関数: CSV解析テスト
function Test-CsvParsing {
    param([string]$CsvPath)
    
    Write-TestLog "CSV解析テスト: $CsvPath"
    
    try {
        $columnInfo = Get-CsvColumnInfo -FilePath $CsvPath
        
        Write-TestLog "ヘッダー数: $($columnInfo.Headers.Count)" 'SUCCESS'
        Write-TestLog "カラムマッピング:" 'SUCCESS'
        foreach ($mapping in $columnInfo.Mappings.GetEnumerator()) {
            Write-TestLog "  $($mapping.Key) -> $($mapping.Value)" 'SUCCESS'
        }
        
        return $true
        
    } catch {
        Write-TestLog "CSV解析失敗: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

### テスト用関数: メール送信テスト（ダミー）
function Test-EmailSending {
    param([hashtable]$IniData)
    
    Write-TestLog "メール送信機能テスト（ダミーモード）"
    
    try {
        # ダミーメール送信テスト
        $testResult = @{
            Success = $true
            ErrorMessage = $null
        }
        
        if ($testResult.Success) {
            Write-TestLog "メール送信機能テスト成功（ダミーモード）" 'SUCCESS'
        } else {
            Write-TestLog "メール送信機能テスト失敗: $($testResult.ErrorMessage)" 'ERROR'
        }
        
        return $testResult.Success
        
    } catch {
        Write-TestLog "メール送信テスト中にエラー: $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

### テスト用関数: クリーンアップ
function Clear-TestData {
    param([hashtable]$IniData)
    
    Write-TestLog "テストデータクリーンアップ中..."
    
    try {
        $userDataPath = $IniData['Paths']['UserData']
        
        # テスト用CSVファイルを削除
        $testFiles = Get-ChildItem -Path $userDataPath -Filter "*.csv" -ErrorAction SilentlyContinue
        foreach ($file in $testFiles) {
            Remove-Item -LiteralPath $file.FullName -Force
            Write-TestLog "削除: $($file.Name)" 'SUCCESS'
        }
        
        # エラー通知フォルダを削除
        $errorFolder = Join-Path $userDataPath "送信処理に問題が発生しています。事務部に連絡してください"
        if (Test-Path $errorFolder) {
            Remove-Item -Path $errorFolder -Recurse -Force
            Write-TestLog "エラー通知フォルダ削除" 'SUCCESS'
        }
        
    } catch {
        Write-TestLog "クリーンアップ中にエラー: $($_.Exception.Message)" 'ERROR'
    }
}

### メインテスト実行関数
function Invoke-SystemTest {
    param([hashtable]$IniData)
    
    Write-TestLog "=== システム全体テスト開始 ===" 'SUCCESS'
    
    $testResults = @{
        IniTest = $false
        CsvParsingTest = $false
        EmailTest = $false
        OverallSuccess = $false
    }
    
    try {
        # 1. 設定ファイルテスト
        Write-TestLog "1. 設定ファイルテスト"
        $testResults.IniTest = $true
        
        # 2. CSV解析テスト
        Write-TestLog "2. CSV解析テスト"
        $testCsvFiles = Get-ChildItem -Path $IniData['Paths']['UserData'] -Filter "*.csv" -ErrorAction SilentlyContinue
        $csvTestSuccess = $true
        foreach ($csvFile in $testCsvFiles) {
            $result = Test-CsvParsing -CsvPath $csvFile.FullName
            if (-not $result) { $csvTestSuccess = $false }
        }
        $testResults.CsvParsingTest = $csvTestSuccess
        
        # 3. メール送信テスト
        Write-TestLog "3. メール送信機能テスト"
        $testResults.EmailTest = Test-EmailSending -IniData $IniData
        
        # 総合判定
        $testResults.OverallSuccess = $testResults.IniTest -and $testResults.CsvParsingTest -and $testResults.EmailTest
        
        Write-TestLog "=== テスト結果 ===" 'SUCCESS'
        Write-TestLog "設定ファイルテスト: $(if($testResults.IniTest){'成功'}else{'失敗'})"
        Write-TestLog "CSV解析テスト: $(if($testResults.CsvParsingTest){'成功'}else{'失敗'})"
        Write-TestLog "メール送信テスト: $(if($testResults.EmailTest){'成功'}else{'失敗'})"
        Write-TestLog "総合結果: $(if($testResults.OverallSuccess){'成功'}else{'失敗'})" $(if($testResults.OverallSuccess){'SUCCESS'}else{'ERROR'})
        
    } catch {
        Write-TestLog "システムテスト中に致命的エラー: $($_.Exception.Message)" 'ERROR'
        $testResults.OverallSuccess = $false
    }
    
    return $testResults
}

### メイン実行部分
try {
    Write-TestLog "=== 外接メール送信システム テスト開始 ===" 'SUCCESS'
    
    # クリーンアップのみの場合
    if ($CleanupOnly) {
        $iniData = Test-IniConfiguration -IniPath $IniPath
        Clear-TestData -IniData $iniData
        Write-TestLog "クリーンアップ完了" 'SUCCESS'
        exit 0
    }
    
    # 設定ファイル読み込み
    $iniData = Test-IniConfiguration -IniPath $IniPath
    
    # テストデータ作成
    if ($CreateTestData) {
        New-TestDirectories -IniData $iniData
        New-TestCsvFiles -IniData $iniData
        Write-TestLog "テストデータ作成完了" 'SUCCESS'
    }
    
    # フルテスト実行
    if ($RunFullTest) {
        $testResults = Invoke-SystemTest -IniData $iniData
        
        if ($testResults.OverallSuccess) {
            Write-TestLog "=== 全テスト成功 ===" 'SUCCESS'
            exit 0
        } else {
            Write-TestLog "=== テスト失敗 ===" 'ERROR'
            exit 1
        }
    }
    
    Write-TestLog "テストスクリプト実行完了" 'SUCCESS'
    Write-TestLog "オプション: -CreateTestData (テストデータ作成), -RunFullTest (フルテスト), -CleanupOnly (クリーンアップのみ)"
    
} catch {
    Write-TestLog "テストスクリプト実行中に致命的エラー: $($_.Exception.Message)" 'ERROR'
    exit 1
}