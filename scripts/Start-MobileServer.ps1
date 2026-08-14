# Serves the mobile dashboard on this PC so your phone can open it on Wi-Fi.
param([int]$Port = 8787)
$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Dash = Join-Path $Root "dashboard"
$Index = Join-Path $Dash "index.html"
if (-not (Test-Path $Index)) {
    & (Join-Path $PSScriptRoot "Process-PmH1.ps1")
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
$ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
    Where-Object { $_.IPAddress -notlike "127.*" -and $_.PrefixOrigin -ne "WellKnown" } |
    Select-Object -ExpandProperty IPAddress -Unique)
foreach ($ip in $ips) {
    try { $listener.Prefixes.Add("http://${ip}:$Port/") } catch {}
}
$listener.Start()
Write-Host "Dashboard running."
Write-Host "This PC:  http://127.0.0.1:$Port/"
foreach ($ip in $ips) { Write-Host "Phone Wi-Fi: http://${ip}:$Port/" }
Write-Host "Keep this window open. On the phone: Safari/Chrome -> Add to Home Screen."
Write-Host "Press Ctrl+C to stop."

$mime = @{
    ".html" = "text/html; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".js"   = "text/javascript; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".png"  = "image/png"
    ".svg"  = "image/svg+xml"
}

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $rel = $ctx.Request.Url.AbsolutePath.TrimStart("/").Replace("/", [IO.Path]::DirectorySeparatorChar)
        if ([string]::IsNullOrWhiteSpace($rel)) { $rel = "index.html" }
        $file = Join-Path $Dash $rel
        if (Test-Path -LiteralPath $file -PathType Leaf) {
            $ext = [IO.Path]::GetExtension($file).ToLowerInvariant()
            $bytes = [IO.File]::ReadAllBytes($file)
            $ctx.Response.StatusCode = 200
            $ctx.Response.ContentType = $(if ($mime.ContainsKey($ext)) { $mime[$ext] } else { "application/octet-stream" })
            $ctx.Response.Headers["Cache-Control"] = "no-store"
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $ctx.Response.StatusCode = 404
        }
        $ctx.Response.Close()
    }
} finally {
    $listener.Stop()
}
