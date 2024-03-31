# デフォルトのOutlookアカウントの送信トレイにアイテムが残ったままの場合に、デスクトップ通知を送信するスクリプト
Add-Type -AssemblyName System.Windows.Forms

# デスクトップ通知を送信する関数
function Show-DesktopNotification ([string]$title, [string]$message) {
    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $notifyIcon.BalloonTipTitle = $title
    $notifyIcon.BalloonTipText = $message
    $notifyIcon.Visible = $true

    $notifyIcon.ShowBalloonTip(5000)  # 通知を表示する時間（ミリ秒）
    Start-Sleep -Seconds 5  # 通知が消えるまでの時間（秒）

    $notifyIcon.Dispose()
}

# Outlookアプリケーションオブジェクトを作成
$outlook = New-Object -ComObject Outlook.Application

# 送信トレイ内のメールアイテム数を取得
$outlookNamespace = $outlook.GetNamespace("MAPI")
$sentItems = $outlookNamespace.GetDefaultFolder(4)
$itemCount = $sentItems.Items.Count

# メールが存在する場合は通知を送る
if ($itemCount -gt 0) {
    $notificationTitle = "Outlook 送信トレイにメールが残っています"
    $notificationMessage = "未送信のメールが $itemCount 件あります。"
    Show-DesktopNotification -title $notificationTitle -message $notificationMessage
} else {
    Write-Host "送信トレイは空です。"
}

# Outlookオブジェクトを解放
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
