# ARASAKA PRINT BRIDGE v1.1
# Commands:
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1 stop
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1 restart
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1 debug
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1 install
#   powershell -ExecutionPolicy Bypass -File .\arasaka-print-bridge.ps1 uninstall
#
# Or just double-click: arasaka-print-bridge.bat

param([string]$Action)


if ((Get-ExecutionPolicy -Scope Process) -ne 'Bypass') {
    $scriptFile = $MyInvocation.MyCommand.Path
    if (-not $scriptFile) { $scriptFile = $PSCommandPath }
    $cmd = "Set-ExecutionPolicy Bypass -Scope Process -Force; & '$($scriptFile -replace "'","''")'"
    if ($Action) { $cmd += " '$Action'" }
    Start-Process powershell -ArgumentList "-NoExit", "-Command", $cmd
    exit 0
}

$Port = 9150
$BridgeUrl = "http://localhost:$Port"
$ScriptPath = $MyInvocation.MyCommand.Path
$ConfigFile = Join-Path $PSScriptRoot "arasaka-print-config.json"
$LogFile = Join-Path $PSScriptRoot "arasaka-print.log"
$TempDir = Join-Path $env:TEMP "arasaka-print"
$SumatraPath = Join-Path $env:LOCALAPPDATA "arasaka-tools\SumatraPDF.exe"
$StartupLink = Join-Path ([Environment]::GetFolderPath('Startup')) "ArasakaPrintBridge.lnk"

if ($Action) {
    switch ($Action.ToLower()) {
        "stop" {
            try {
                $r = Invoke-RestMethod "$BridgeUrl/shutdown" -Method POST -TimeoutSec 3
                Write-Host "Bridge stopped." -ForegroundColor Green
            } catch { Write-Host "Bridge not running." -ForegroundColor Yellow }
            exit 0
        }
        "restart" {
            try { Invoke-RestMethod "$BridgeUrl/shutdown" -Method POST -TimeoutSec 3 | Out-Null } catch { }
            Start-Sleep -Seconds 2
            $cmd = "Set-ExecutionPolicy Bypass -Scope Process -Force; & '$($ScriptPath -replace "'","''")'"
            Start-Process powershell -ArgumentList "-NoExit", "-Command", $cmd
            Write-Host "Bridge restarted." -ForegroundColor Green
            exit 0
        }
        "debug" {
            if (Test-Path $LogFile) {
                Get-Content $LogFile -Tail 50
            } else { Write-Host "No log file found." -ForegroundColor Yellow }
            exit 0
        }
        "install" {
            $ws = New-Object -ComObject WScript.Shell
            $sc = $ws.CreateShortcut($StartupLink)
            $sc.TargetPath = "powershell.exe"
            $escapedPath = $ScriptPath -replace "'", "''"
            $sc.Arguments = "-NoExit -Command `"Set-ExecutionPolicy Bypass -Scope Process -Force; & '$escapedPath'`""
            $sc.WorkingDirectory = $PSScriptRoot
            $sc.WindowStyle = 7
            $sc.Save()
            Write-Host "Autostart installed: $StartupLink" -ForegroundColor Green
            exit 0
        }
        "uninstall" {
            if (Test-Path $StartupLink) {
                Remove-Item $StartupLink -Force
                Write-Host "Autostart removed." -ForegroundColor Green
            } else { Write-Host "No autostart found." -ForegroundColor Yellow }
            exit 0
        }
        default { Write-Host "Unknown action: $Action" -ForegroundColor Red; exit 1 }
    }
}

function Log($msg) {
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"
    Write-Host "  $line"
    Add-Content $LogFile $line -ErrorAction SilentlyContinue
}

function Load-Config {
    if (Test-Path $ConfigFile) {
        try { return (Get-Content $ConfigFile -Raw | ConvertFrom-Json) } catch { }
    }
    return [PSCustomObject]@{ printer = "" }
}

function Save-Config($cfg) { $cfg | ConvertTo-Json | Set-Content $ConfigFile -Encoding UTF8 }

if (-not (Test-Path $TempDir)) { New-Item $TempDir -ItemType Directory -Force | Out-Null }

function Ensure-SumatraPDF {
    if (Test-Path $script:SumatraPath) { return $true }

    $searchPaths = @(
        (Join-Path $env:LOCALAPPDATA "SumatraPDF\SumatraPDF.exe"),
        (Join-Path ${env:ProgramFiles} "SumatraPDF\SumatraPDF.exe")
    )
    if ($env:ProgramW6432) { $searchPaths += Join-Path $env:ProgramW6432 "SumatraPDF\SumatraPDF.exe" }
    if (${env:ProgramFiles(x86)}) { $searchPaths += Join-Path ${env:ProgramFiles(x86)} "SumatraPDF\SumatraPDF.exe" }
    $searchPaths += Join-Path $PSScriptRoot "SumatraPDF.exe"

    foreach ($sp in $searchPaths) {
        if ($sp -and (Test-Path $sp)) {
            $script:SumatraPath = $sp
            Write-Host "  SumatraPDF found: $sp" -ForegroundColor Green
            return $true
        }
    }

    $inPath = Get-Command SumatraPDF -ErrorAction SilentlyContinue
    if ($inPath) {
        $script:SumatraPath = $inPath.Source
        Write-Host "  SumatraPDF found in PATH: $($inPath.Source)" -ForegroundColor Green
        return $true
    }

    Write-Host "  SumatraPDF not found. Downloading portable version..." -ForegroundColor Yellow
    $toolsDir = Split-Path $script:SumatraPath -Parent
    if (-not (Test-Path $toolsDir)) { New-Item $toolsDir -ItemType Directory -Force | Out-Null }

    $dlUrl = "https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.exe"
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($dlUrl, $script:SumatraPath)
        $wc.Dispose()
        if (Test-Path $script:SumatraPath) {
            Write-Host "  SumatraPDF downloaded OK -> $($script:SumatraPath)" -ForegroundColor Green
            return $true
        }
    } catch {
        Write-Host "  Auto-download failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host "  ERROR: SumatraPDF could not be found or downloaded." -ForegroundColor Red
    Write-Host "  Install it manually from https://www.sumatrapdfreader.org/download-free-pdf-viewer" -ForegroundColor Yellow
    return $false
}

function Get-PrinterList {
    try { @(Get-CimInstance Win32_Printer | Select-Object -ExpandProperty Name) } catch { @() }
}

function Download-Pdf($url) {
    $fp = Join-Path $TempDir "arasaka_$(Get-Date -Format 'yyyyMMdd_HHmmss').pdf"
    try {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($url, $fp)
        $wc.Dispose()
        return $fp
    } catch { return $null }
}

function Print-Pdf($filePath, $printerName, $copies) {
    if (-not $copies) { $copies = 1 }
    if (-not (Test-Path $SumatraPath)) { return "SumatraPDF not found at $SumatraPath" }
    for ($i = 0; $i -lt $copies; $i++) {
        $p = Start-Process -FilePath $SumatraPath -ArgumentList "-print-to `"$printerName`" -print-settings `"simplex`" -silent `"$filePath`"" -PassThru -WindowStyle Hidden
        $null = $p.WaitForExit(30000)
        if (-not $p.HasExited) { $p.Kill() }
    }
    Start-Job -ScriptBlock { param($f); Start-Sleep 15; Remove-Item $f -Force -EA SilentlyContinue } -ArgumentList $filePath | Out-Null
    return $null
}

function Send-Response($stream, $code, $status, $body) {
    $b = [System.Text.Encoding]::UTF8.GetBytes($body)
    $h = "HTTP/1.1 $code $status`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($b.Length)`r`nAccess-Control-Allow-Origin: *`r`nAccess-Control-Allow-Methods: GET, POST, OPTIONS`r`nAccess-Control-Allow-Headers: Content-Type`r`nConnection: close`r`n`r`n"
    $hb = [System.Text.Encoding]::ASCII.GetBytes($h)
    $stream.Write($hb, 0, $hb.Length); $stream.Write($b, 0, $b.Length); $stream.Flush()
}

function Ok($s, $o) { Send-Response $s 200 "OK" ($o | ConvertTo-Json -Compress) }
function Err($s, $c, $m) { Send-Response $s $c "Error" (@{success=$false;error=$m} | ConvertTo-Json -Compress) }

function Parse-Request($client) {
    $s = $client.GetStream(); $s.ReadTimeout = 5000
    $buf = New-Object byte[] 65536; $raw = ""
    try { do { $n = $s.Read($buf,0,$buf.Length); if($n -gt 0){ $raw += [System.Text.Encoding]::UTF8.GetString($buf,0,$n) } } while($s.DataAvailable) } catch {}
    if (-not $raw) { return $null }
    $parts = ($raw -split "`r`n")[0] -split " "
    $bi = $raw.IndexOf("`r`n`r`n"); $body = ""; if($bi -ge 0){ $body = $raw.Substring($bi+4) }
    return @{ Method=$parts[0]; Path=$parts[1]; Body=$body; Stream=$s; Client=$client }
}

trap { Log "FATAL: $($_.Exception.Message)"; Read-Host "Enter to exit"; exit 1 }

if (Test-Path $LogFile) { Clear-Content $LogFile -ErrorAction SilentlyContinue }

$config = Load-Config
$running = $true

Write-Host ""
Write-Host "  ARASAKA PRINT BRIDGE v1.1" -ForegroundColor Cyan
Write-Host "  Port $Port | Chrome headless" -ForegroundColor Cyan

$allPrinters = Get-PrinterList

if (-not $config.printer -or ($allPrinters -notcontains $config.printer)) {
    Write-Host ""
    Write-Host "  Kein Drucker konfiguriert. Bitte waehlen:" -ForegroundColor Yellow
    Write-Host ""
    for ($i = 0; $i -lt $allPrinters.Count; $i++) {
        Write-Host "    [$($i+1)] $($allPrinters[$i])" -ForegroundColor White
    }
    Write-Host ""
    $choice = 0
    while ($choice -lt 1 -or $choice -gt $allPrinters.Count) {
        $input = Read-Host "  Nummer eingeben (1-$($allPrinters.Count))"
        try { $choice = [int]$input } catch { $choice = 0 }
    }
    $config.printer = $allPrinters[$choice - 1]
    Save-Config $config
    Write-Host ""
    Write-Host "  Drucker gesetzt: $($config.printer)" -ForegroundColor Green
}

if (-not (Ensure-SumatraPDF)) {
    Write-Host "" 
    Write-Host "  Bridge cannot start without SumatraPDF." -ForegroundColor Red
    Read-Host "  Enter to exit"
    exit 1
}
Write-Host "  SumatraPDF: $SumatraPath" -ForegroundColor DarkGray

Write-Host "  Printer: $($config.printer)" -ForegroundColor Green
Write-Host "  Printers:" -ForegroundColor DarkGray
foreach ($p in $allPrinters) {
    if ($p -eq $config.printer) { Write-Host "    > $p" -ForegroundColor Green }
    else { Write-Host "    - $p" -ForegroundColor DarkGray }
}
Write-Host "  Listening..." -ForegroundColor Green
Write-Host ""

try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
    $listener.Start()
} catch {
    Log "Cannot bind port $Port - $($_.Exception.Message)"
    Read-Host "Enter to exit"; exit 1
}

try {
    while ($running) {
        if (-not $listener.Pending()) { Start-Sleep -Milliseconds 100; continue }
        $client = $listener.AcceptTcpClient()
        $req = Parse-Request $client
        if (-not $req) { $client.Close(); continue }
        $m = $req.Method; $pa = $req.Path; $st = $req.Stream
        try {
            if ($m -eq "OPTIONS") { Send-Response $st 204 "No Content" "" }
            elseif ($m -eq "GET" -and $pa -eq "/status") {
                Log "GET /status"
                Ok $st @{ success=$true; service="arasaka-print-bridge"; version="1.1"; printer=$config.printer; configured=[bool]$config.printer }
            }
            elseif ($m -eq "GET" -and $pa -eq "/printers") {
                Log "GET /printers"
                Ok $st @{ success=$true; printers=(Get-PrinterList); selected=$config.printer }
            }
            elseif ($m -eq "POST" -and $pa -eq "/config") {
                try { $d = $req.Body | ConvertFrom-Json; $np = $d.printer } catch { Err $st 400 "Bad JSON"; $client.Close(); continue }
                if (-not $np) { Err $st 400 "Missing printer"; $client.Close(); continue }
                if ((Get-PrinterList) -notcontains $np) { Err $st 400 "Printer not found"; $client.Close(); continue }
                $config.printer = $np; Save-Config $config
                Log "Printer set: $np"
                Ok $st @{ success=$true; printer=$np }
            }
            elseif ($m -eq "POST" -and $pa -eq "/print") {
                if (-not $config.printer) { Err $st 400 "No printer set"; $client.Close(); continue }
                try { $d = $req.Body | ConvertFrom-Json } catch { Err $st 400 "Bad JSON"; $client.Close(); continue }
                $url = $d.url; $cp = 1; if($d.copies){ $cp = [int]$d.copies }
                if (-not $url) { Err $st 400 "Missing url"; $client.Close(); continue }
                Log "PRINT $url x$cp"
                $fp = Download-Pdf $url
                if (-not $fp) { Log "Download FAIL"; Err $st 500 "Download failed"; $client.Close(); continue }
                $pe = Print-Pdf $fp $config.printer $cp
                if ($pe) { Log "Print FAIL: $pe"; Err $st 500 "Print failed: $pe" }
                else { Log "Print OK -> $($config.printer)"; Ok $st @{ success=$true; printer=$config.printer; copies=$cp; message="Sent" } }
            }
            elseif ($m -eq "POST" -and $pa -eq "/shutdown") {
                Log "Shutdown"; Ok $st @{ success=$true; message="Bye" }; $running = $false
            }
            else { Err $st 404 "Not found" }
        } catch {
            Log "ERR: $($_.Exception.Message)"
            try { Err $st 500 $_.Exception.Message } catch {}
        } finally { try { $client.Close() } catch {} }
    }
} finally {
    $listener.Stop()
    Log "Stopped."
}
