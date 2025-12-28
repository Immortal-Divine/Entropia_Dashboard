<# extras.psm1 #>

$ProxyUrl = "https://worldboss.entropia-dashboard.workers.dev/"

function Send-Message {
    param(
        $code,
        $bossName
    )

    $boss = $global:DashboardConfig.Resources.BossData[$bossName]

    if (-not $boss) {
        [Windows.Forms.MessageBox]::Show("Unknown boss: $bossName", "Error")
        return
    }

    $bossImageUrl = $boss.url
    $bossRolePing = $boss.role
    $bossDisplay  = $boss.name

    $embed = @{
        title       = "$bossDisplay spawned!"
        color       = 16724556
        description = " "
        footer      = @{ text = "Sent via Entropia Dashboard`nhttps://immortal-divine.github.io/Entropia_Dashboard/" }
    }

    if (-not [string]::IsNullOrEmpty($bossImageUrl)) {
        $embed.image = @{ url = $bossImageUrl }
    }

    $discordPayload = @{
        content = "<@&$bossRolePing> **$bossDisplay** spawned!"
        embeds  = @($embed)
    }

    $finalBody = @{
        auth_code = $code
        message   = $discordPayload
    } | ConvertTo-Json -Depth 10

    try {
        Invoke-RestMethod -Uri $ProxyUrl -Method Post -Body $finalBody -ContentType "application/json" -ErrorAction Stop
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $details = $reader.ReadToEnd()
            $errorMsg = "$errorMsg`nServer Response: $details"
        }
        		Show-DarkMessageBox $global:DashboardConfig.UI.WorldbossForm "Failed to send:`n$errorMsg" "Can't send ping" 'Ok' 'Error'
    }
}