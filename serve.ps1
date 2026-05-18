$port = 3457

# Kill any existing process on this port
$existing = netstat -ano | Select-String ":${port}\s" | ForEach-Object {
    ($_ -split '\s+')[-1]
} | Sort-Object -Unique | Where-Object { $_ -ne "0" -and $_ -ne "4" }
if ($existing) {
    foreach ($p in $existing) {
        try { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue; Write-Host "[INFO] Killed old process PID $p" } catch {}
    }
    Start-Sleep -Milliseconds 500
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:${port}/")
try {
    $listener.Prefixes.Add("http://+:${port}/")
    $listener.Start()
} catch {
    Write-Host "[WARN] Trying localhost only..." -ForegroundColor Yellow
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:${port}/")
    $listener.Start()
}

$ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -ne "127.0.0.1" -and $_.PrefixOrigin -ne "WellKnown" } | Select-Object -First 1).IPAddress
Write-Host ""
Write-Host "=== FFB Ramp Server ===" -ForegroundColor Green
Write-Host "PC:        http://localhost:${port}/" -ForegroundColor Cyan
if ($ip) { Write-Host "Mobile:    http://${ip}:${port}/" -ForegroundColor Cyan }
Write-Host ""

# Ensure data folder exists
$dataDir = Join-Path $PSScriptRoot "data"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
Write-Host "Data folder: $dataDir"

function Send-Json($response, $obj, $status) {
    if (-not $status) { $status = 200 }
    $response.StatusCode = $status
    $response.ContentType = "application/json; charset=utf-8"
    $response.Headers.Add("Access-Control-Allow-Origin", "*")
    $json = $obj | ConvertTo-Json -Depth 20 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
}

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    $path = $request.Url.LocalPath
    $method = $request.HttpMethod

    # CORS preflight
    if ($method -eq "OPTIONS") {
        $response.StatusCode = 204
        $response.Headers.Add("Access-Control-Allow-Origin", "*")
        $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
        $response.Close()
        continue
    }

    try {
        # ===== API: GET /api/dates =====
        if ($path -eq "/api/dates" -and $method -eq "GET") {
            $files = Get-ChildItem -Path $dataDir -Filter "*.json" -ErrorAction SilentlyContinue |
                     Sort-Object Name |
                     ForEach-Object { $_.BaseName }
            Send-Json $response @{ dates = @($files) }
        }

        # ===== API: GET /api/data/{date} =====
        elseif ($path -match "^/api/data/(\d{4}-\d{2}-\d{2})$" -and $method -eq "GET") {
            $date = $Matches[1]
            $file = Join-Path $dataDir "$date.json"
            if (Test-Path $file) {
                $raw = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
                $response.StatusCode = 200
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
                $response.ContentLength64 = $bytes.Length
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
            } else {
                Send-Json $response @{ error = "not_found"; date = $date } 404
            }
        }

        # ===== API: POST /api/data/{date} =====
        elseif ($path -match "^/api/data/(\d{4}-\d{2}-\d{2})$" -and $method -eq "POST") {
            $date = $Matches[1]
            $file = Join-Path $dataDir "$date.json"
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $body = $reader.ReadToEnd()
            $reader.Close()
            [System.IO.File]::WriteAllText($file, $body, [System.Text.Encoding]::UTF8)
            Write-Host "[SAVE] $date.json ($(($body.Length / 1024).ToString('F1')) KB)"
            Send-Json $response @{ ok = $true; date = $date; size = $body.Length }
        }

        # ===== API: DELETE /api/data/{date} =====
        elseif ($path -match "^/api/data/(\d{4}-\d{2}-\d{2})$" -and $method -eq "DELETE") {
            $date = $Matches[1]
            $file = Join-Path $dataDir "$date.json"
            if (Test-Path $file) {
                Remove-Item $file -Force
                Write-Host "[DELETE] $date.json"
                Send-Json $response @{ ok = $true; deleted = $date }
            } else {
                Send-Json $response @{ error = "not_found" } 404
            }
        }

        # ===== API: GET /api/export/{date} — Excel HTML =====
        elseif ($path -match "^/api/export/(\d{4}-\d{2}-\d{2})$" -and $method -eq "GET") {
            $date = $Matches[1]
            $file = Join-Path $dataDir "$date.json"
            if (Test-Path $file) {
                # Just return the JSON, the client will build Excel
                $raw = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
                $response.StatusCode = 200
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
                $response.ContentLength64 = $bytes.Length
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
            } else {
                Send-Json $response @{ error = "not_found" } 404
            }
        }

        # ===== STATIC FILES =====
        else {
            if ($path -eq "/" -or $path -eq "") { $path = "/index.html" }
            $filePath = Join-Path $PSScriptRoot $path.TrimStart("/")
            if (Test-Path $filePath) {
                $content = [System.IO.File]::ReadAllBytes($filePath)
                $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
                $mime = switch ($ext) {
                    ".html" { "text/html; charset=utf-8" }
                    ".js"   { "application/javascript" }
                    ".css"  { "text/css" }
                    ".png"  { "image/png" }
                    ".jpg"  { "image/jpeg" }
                    ".svg"  { "image/svg+xml" }
                    ".json" { "application/json; charset=utf-8" }
                    ".xls"  { "application/vnd.ms-excel" }
                    default { "application/octet-stream" }
                }
                $response.ContentType = $mime
                $response.ContentLength64 = $content.Length
                $response.OutputStream.Write($content, 0, $content.Length)
            } else {
                $response.StatusCode = 404
                $msg = [System.Text.Encoding]::UTF8.GetBytes("Not Found: $path")
                $response.OutputStream.Write($msg, 0, $msg.Length)
            }
        }
    }
    catch {
        Write-Host "[ERROR] $_"
        try {
            $response.StatusCode = 500
            $errMsg = [System.Text.Encoding]::UTF8.GetBytes("Server Error")
            $response.OutputStream.Write($errMsg, 0, $errMsg.Length)
        } catch {}
    }

    $response.Close()
}
