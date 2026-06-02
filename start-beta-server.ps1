# start-beta-server.ps1 - Levanta la beta casera: ngrok (relay del game) + server.exe acotado.
#
# Hace, en un solo paso:
#   1) Se auto-eleva a Administrador (server.exe lo necesita: si no, fallan migrations/ODBC/OLE DB).
#   2) Abre ngrok TCP del puerto del game (default 7667, region sa) en su propia ventana
#      -> de ahi sale el host:puerto publico (0.tcp.sa.ngrok.io:XXXXX) que va en el login del cliente.
#   3) Lanza server.exe con prioridad baja (default BelowNormal) y afinidad limitada (default 2 cores),
#      para que no mate la maquina/VM.
#
# NOTA: la palanca de CPU mas fuerte es el BUILD con UNLOCK_CPU=0 (loop hace Sleep(1); ya compilado).
# Este script ademas acota el runtime (prioridad + afinidad).
#
# Uso (PowerShell normal; pide UAC solo):
#   powershell -ExecutionPolicy Bypass -File C:\AO20\dev\server\start-beta-server.ps1
#   ...-File ... -Priority Idle -Cores 1            # lo mas conservador
#   ...-File ... -Region us -GamePort 7667          # otra region de ngrok
#   ...-File ... -SkipNgrok                          # si ngrok ya lo levantas aparte
param(
  [ValidateSet('Idle','BelowNormal','Normal')]
  [string]$Priority  = 'BelowNormal',
  [int]$Cores        = 2,
  [string]$ServerDir = 'C:\AO20\dev\server',
  [string]$Exe       = 'server.exe',
  [int]$GamePort     = 7667,
  [string]$Region    = 'sa',
  [string]$NgrokExe  = 'ngrok',
  [switch]$SkipNgrok
)

$exePath = Join-Path $ServerDir $Exe
if (-not (Test-Path -LiteralPath $exePath)) {
  Write-Host "[ERROR] No existe $exePath. Compila el server primero (IDE VB6, build_tmp.bat = UNLOCK_CPU=0)." -ForegroundColor Red
  exit 1
}

# --- Auto-elevacion ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
  Write-Host "[info] Re-lanzando como Administrador (UAC)..." -ForegroundColor Yellow
  $a = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $PSCommandPath),
         '-Priority',$Priority,'-Cores',$Cores,'-ServerDir',('"{0}"' -f $ServerDir),'-Exe',('"{0}"' -f $Exe),
         '-GamePort',$GamePort,'-Region',$Region,'-NgrokExe',('"{0}"' -f $NgrokExe))
  if ($SkipNgrok) { $a += '-SkipNgrok' }
  Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $a
  exit 0
}

# --- 1) ngrok TCP del game ---
if (-not $SkipNgrok) {
  $ng = (Get-Command $NgrokExe -ErrorAction SilentlyContinue).Source
  if (-not $ng) { $ng = $NgrokExe }
  if ((Test-Path -LiteralPath $ng) -or (Get-Command $NgrokExe -ErrorAction SilentlyContinue)) {
    Write-Host ("[ngrok] tcp {0} --region {1}  (anota el host:puerto que aparezca)" -f $GamePort, $Region) -ForegroundColor Cyan
    Start-Process -FilePath $ng -ArgumentList ('tcp {0} --region {1}' -f $GamePort, $Region)
  } else {
    Write-Host "[AVISO] No encontre ngrok ('$NgrokExe'). Pasa -NgrokExe 'C:\ruta\ngrok.exe' o levantalo a mano." -ForegroundColor Yellow
  }
}

# --- 2) server.exe acotado (prioridad + afinidad) ---
$cpu  = [Environment]::ProcessorCount
if ($Cores -lt 1)    { $Cores = 1 }
if ($Cores -gt $cpu) { $Cores = $cpu }
$mask = ([int]([Math]::Pow(2, $Cores)) - 1)

Write-Host ("[server] {0}  | afinidad {1} core(s) (mask 0x{2:X})  prioridad {3}" -f $exePath, $Cores, $mask, $Priority) -ForegroundColor Cyan
$p = Start-Process -FilePath $exePath -WorkingDirectory $ServerDir -PassThru

$set = $false
for ($i = 0; $i -lt 10 -and -not $set; $i++) {
  try {
    if ($p.HasExited) { break }
    $p.ProcessorAffinity = [IntPtr]$mask
    $p.PriorityClass     = [System.Diagnostics.ProcessPriorityClass]::$Priority
    $set = $true
  } catch { Start-Sleep -Milliseconds 200; $p.Refresh() }
}

if ($set) {
  Write-Host ("[ok] server PID {0} (afinidad {1} core, prioridad {2}). Espera 'Escuchando conexiones entrantes ...'." -f $p.Id, $Cores, $Priority) -ForegroundColor Green
  if (-not $SkipNgrok) { Write-Host "[ok] Pasa el host:puerto de la ventana de ngrok al login del cliente (Servidor/Puerto)." -ForegroundColor Green }
} else {
  Write-Host "[AVISO] No pude fijar afinidad/prioridad (proceso salio o nego handle). Revisa la ventana del server." -ForegroundColor Yellow
}
