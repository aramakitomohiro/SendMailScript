#===========================
#BulkSendingExternalEmail.ps1
#外接メール送信_メール送信実行（本番用）
#起動契機:
# - JP1から5分間隔で起動し、CSVファイルがあれば処理開始
#処理結果の判定方法:
# - 0=成功  1=処理なし  2=一部失敗  3=致命的エラー
#===========================

param(
  [string]$IniPath = "$PSScriptRoot\..\001_InspectingAndMovingInputData\mailsettings_本番用.ini",
  [switch]$VerboseLog
)

# 共通モジュールをインポート
Import-Module "$PSScriptRoot\BulkSendingExternalEmail.psm1" -Force

# UTF-8エンコーディングを設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# VerboseLogスイッチを全体で利用可能にする
$script:VerboseLogEnabled = $VerboseLog.IsPresent
    
### B-0: ログ出力用関数（INFO/WARN/ERROR/VERBOSEレベルでログを出力）
function Write-Log {
  param([string]$Message, [ValidateSet('INFO','WARN','ERROR','VERBOSE')][string]$Level='INFO')
  
  # VERBOSEレベルの場合、VerboseLogスイッチがない時はスキップ
  if ($Level -eq 'VERBOSE' -and -not $script:VerboseLogEnabled) { return }
  
  $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
  $logLine = "$timestamp [$Level] $Message"
  Write-Host $logLine
  try { 
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $logLine -Encoding UTF8 }
    if ($script:MultiScheduleLogFile) { Add-Content -LiteralPath $script:MultiScheduleLogFile -Value $logLine -Encoding UTF8 }
  } catch {}
}

### B-1: 肥大化ログファイル削除判定
function Remove-OldLogFiles {
  param([string]$TargetPath, [switch]$WhatIf)
  
  # エラー変数をクリア
  $Error.Clear()
  
  # 保存期間　日単位で指定
  $Days = 731
  
  try {
    # 指定日以前に更新されたログファイル＝削除対象
    $target = Get-ChildItem $TargetPath -Filter "*.log" -ErrorAction SilentlyContinue | Where-Object {$_.CreationTime -lt (Get-Date).AddDays(-1 * $Days) }
    
    If ($Error.Count -gt 0){
      Write-Log "ログファイル削除処理でエラー: $($Error[0])" 'ERROR'
      return $false
    }
  } catch {
    Write-Log "ログファイル削除処理でエラー: $($_.Exception.Message)" 'ERROR'
    return $false
  }
  
  $target | ForEach-Object {
    try {
      if ($WhatIf) {
        Write-Log "[What-If] 削除対象: $($_.FullName)" 'INFO'
      } else {
        Write-Log "削除対象: $($_.FullName)" 'INFO'
        # 削除
        Remove-Item $_.FullName -Force -ErrorAction Stop
        Write-Log "ログファイル削除完了: $($_.FullName)" 'INFO'
      }
    } catch {
      Write-Log "ログファイル削除失敗: $($_.FullName) - $($_.Exception.Message)" 'ERROR'
      return $false
    }
  }
  return $true
}

### B-2: 失敗時のCSVファイル処理（UserDataへ移動のみ）
function Handle-FailedCsv {
  param([System.IO.FileInfo]$CsvFile, [string]$UserData, [string]$ErrorReason)
  
  try {
    # エラー通知フォルダを作成
    $errorNotificationFolder = Join-Path $UserData "送信処理に問題が発生しています。事務部に連絡してください"
    if (-not (Test-Path $errorNotificationFolder)) {
      New-Item -ItemType Directory -Path $errorNotificationFolder -Force | Out-Null
      Write-Log "エラー通知フォルダ作成: $errorNotificationFolder" 'ERROR'
    }
    
    # CSVファイルをUserDataに戻す
    $returnFileName = "【送信失敗_$(Get-Date -Format 'yyyy-MM-dd-HH-mm')】_$($CsvFile.Name)"
    $returnPath = Join-Path $UserData $returnFileName
    
    # ファイル内容をUTF-8 BOM付きで保存
    $content = Get-Content -LiteralPath $CsvFile.FullName -Encoding UTF8 -Raw
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($returnPath, $content, $utf8WithBom)
    
    # エラー詳細ファイルを作成
    $errorDetailFile = Join-Path $errorNotificationFolder "エラー詳細_$(Get-Date -Format 'yyyy-MM-dd-HH-mm').txt"
    $errorDetail = @"
送信失敗ファイル: $($CsvFile.Name)
失敗理由: $ErrorReason
失敗日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
返却ファイル: $returnFileName
"@
    [System.IO.File]::WriteAllText($errorDetailFile, $errorDetail, [System.Text.Encoding]::UTF8)
    
    Write-Log "CSVファイル送信失敗処理完了: $returnPath" 'ERROR'
    Write-Log "エラー詳細ファイル作成: $errorDetailFile" 'ERROR'
    
    # 元のファイルを削除
    Remove-Item -LiteralPath $CsvFile.FullName -Force
    
  } catch {
    Write-Log "失敗時のCSVファイル処理でエラー: $($_.Exception.Message)" 'ERROR'
    throw
  }
}

### B-3: メイン処理関数
function Invoke-BulkEmailSending {
  param([hashtable]$IniData)
  
  $pathSettings = $IniData['Paths']
  $smtpSettings = $IniData['SMTP']
  $generalSettings = $IniData['General']
  
  $InspectingAndMovingInputData = $pathSettings['001_InspectingAndMovingInputData']
  $BulkSendingExternalEmails = $pathSettings['002_BulkSendingExternalEmails']
  $UserData = $pathSettings['UserData']
  
  # 送信準備完了ファイルを検索
  $readyFiles = Get-ChildItem -LiteralPath $InspectingAndMovingInputData -Filter "【送信準備完了_*】_*.csv" -File -ErrorAction SilentlyContinue
  
  if (-not $readyFiles) {
    Write-Log "送信準備完了ファイルなし。処理終了" 'INFO'
    return 1  # 処理なし
  }
  
  $totalFiles = $readyFiles.Count
  $successCount = 0
  $failureCount = 0
  
  Write-Log "送信準備完了ファイル発見: $totalFiles 件" 'INFO'
  
  foreach ($csvFile in $readyFiles) {
    try {
      Write-Log "処理開始: $($csvFile.Name)" 'INFO'
      
      # CSVファイルを送信処理領域に移動
      $processingFileName = $csvFile.Name -replace '【送信準備完了_.*?】_', '【送信処理中_' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '】_'
      $processingPath = Join-Path $BulkSendingExternalEmails $processingFileName
      
      # ファイル移動（UTF-8 BOM付きを維持）
      $content = Get-Content -LiteralPath $csvFile.FullName -Encoding UTF8 -Raw
      $utf8WithBom = New-Object System.Text.UTF8Encoding $true
      [System.IO.File]::WriteAllText($processingPath, $content, $utf8WithBom)
      Remove-Item -LiteralPath $csvFile.FullName -Force
      
      # CSVファイル処理
      $processingFile = Get-Item $processingPath
      $result = Send-CsvEmails -CsvFile $processingFile -IniData $IniData
      
      if ($result.Success) {
        $successCount++
        Write-Log "送信完了: $($csvFile.Name)" 'INFO'
        
        # 成功時は完了ファイルとして保存
        $completedFileName = $processingFileName -replace '【送信処理中_.*?】_', '【送信完了_' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '】_'
        $completedPath = Join-Path $BulkSendingExternalEmails $completedFileName
        Move-Item -LiteralPath $processingPath -Destination $completedPath -Force
        
      } else {
        $failureCount++
        Write-Log "送信失敗: $($csvFile.Name) - $($result.ErrorMessage)" 'ERROR'
        Handle-FailedCsv -CsvFile $processingFile -UserData $UserData -ErrorReason $result.ErrorMessage
      }
      
    } catch {
      $failureCount++
      $errorMessage = "ファイル処理中に致命的エラー: $($_.Exception.Message)"
      Write-Log $errorMessage 'ERROR'
      
      try {
        if (Test-Path $processingPath) {
          $processingFile = Get-Item $processingPath
          Handle-FailedCsv -CsvFile $processingFile -UserData $UserData -ErrorReason $errorMessage
        }
      } catch {
        Write-Log "エラーハンドリング中にさらなるエラー: $($_.Exception.Message)" 'ERROR'
      }
    }
  }
  
  Write-Log "バッチ処理完了: 成功=$successCount, 失敗=$failureCount, 合計=$totalFiles" 'INFO'
  
  # 戻り値の決定
  if ($failureCount -eq 0) {
    return 0  # 成功
  } elseif ($successCount -gt 0) {
    return 2  # 一部失敗
  } else {
    return 3  # 致命的エラー
  }
}

### メイン実行部分
try {
  # INI設定を読み込み
  $iniData = Get-IniData -Path $IniPath
  $pathSettings = $iniData['Paths']
  
  # ログファイル設定
  $logDirectory = $pathSettings['002_BulkSendingExternalEmailsLog']
  if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
  }
  
  $script:LogFile = Join-Path $logDirectory ("B_" + (Get-Date -Format 'yyyy_MM_dd') + ".log")
  
  Write-Log "Start ScriptB (BulkSendingExternalEmail)" 'INFO'
  
  # 古いログファイルの削除
  $logCleanupResult = Remove-OldLogFiles -TargetPath $logDirectory
  if (-not $logCleanupResult) {
    Write-Log "ログファイル削除処理に問題がありましたが、続行します" 'WARN'
  }
  
  # メイン処理実行
  $exitCode = Invoke-BulkEmailSending -IniData $iniData
  
  Write-Log "ScriptB完了. ExitCode=$exitCode" 'INFO'
  exit $exitCode
  
} catch {
  $errorMessage = "ScriptBで致命的エラー: $($_.Exception.Message)"
  Write-Log $errorMessage 'ERROR'
  Write-Log "スタックトレース: $($_.ScriptStackTrace)" 'ERROR'
  exit 3  # 致命的エラー
}