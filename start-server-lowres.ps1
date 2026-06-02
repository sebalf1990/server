# start-server-lowres.ps1 - Lanza el server AO20 acotando su footprint en runtime.
#
# Que hace:
#   - Se auto-eleva a Administrador (server.exe lo necesita: si no, fallan migrations/ODBC/OLE DB).
#   - Lanza server.exe con el working dir correcto (dev\server, lee ..\Recursos de forma relativa).
#   - Le baja la PRIORIDAD (default BelowNormal) y le limita la AFINIDAD de CPU (default 2 cores),
#     para que no se coma la maquina de casa mientras corren ngrok / el web / etc.
#
# OJO - palanca de CPU mas fuerte = el BUILD, no esto:
#   Con UNLOCK_CPU=1 (build canonico) el loop principal NO duerme -> clava 1 core al 100%
#   (General.bas: '#If UNLOCK_CPU = 0 Then Call Sleep(1)'). Para CPU baja de verdad, compila
#   con UNLOCK_CPU=0 (ya existe dev\server\build_tmp.bat). Este launcher solo acota; con el build
#   UNLOCK_CPU=1 la afinidad evita que tome TODOS los cores, pero igual gira los que le dejes.
#
# Uso (PowerShell normal; el script pide UAC solo):
#   powershell -ExecutionPolicy Bypass -File C:\AO20\dev\server\start-server-lowres.ps1
#   ...-File ... -Priority Idle -Cores 1        # lo mas conservador
#   ...-File ... -Priority Normal -Cores 4      # mas holgado
param(
  [ValidateSet('Idle','BelowNormal','Normal')]
  [string]$Priority  = 'BelowNormal',
  [int]$Cores        = 2,
  [string]$ServerDir = 'C:\AO20\dev\server',
  [string]$Exe       = 'server.exe'
)

$exePath = Join-Path $ServerDir $Exe
if (-not (Test-Path -LiteralPath $exePath)) {
  Write-Host "[ERROR] No existe $exePath. Compila el server primero (IDE VB6)." -ForegroundColor Red
  exit 1
}

# --- Auto-elevacion: server.exe necesita token de Administrador ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
  Write-Host "[info] Re-lanzando como Administrador (UAC)..." -ForegroundColor Yellow
  $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $PSCommandPath),
               '-Priority',$Priority,'-Cores',$Cores,'-ServerDir',('"{0}"' -f $ServerDir),'-Exe',('"{0}"' -f $Exe))
  Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList
  exit 0
}

# --- Ya elevado: lanzar server.exe con su CWD y ajustar prioridad + afinidad ---
$cpu  = [Environment]::ProcessorCount
if ($Cores -lt 1)   { $Cores = 1 }
if ($Cores -gt $cpu){ $Cores = $cpu }
$mask = ([int]([Math]::Pow(2, $Cores)) - 1)   # los $Cores cores mas bajos

Write-Host "[start] $exePath" -ForegroundColor Cyan
Write-Host ("[start] CPUs={0}  ->  afinidad {1} core(s) (mask 0x{2:X})  prioridad {3}" -f $cpu, $Cores, $mask, $Priority) -ForegroundColor Cyan

$p = Start-Process -FilePath $exePath -WorkingDirectory $ServerDir -PassThru

# El proceso recien arranca; reintentar el set un par de veces por si todavia no acepta el handle.
$set = $false
for ($i = 0; $i -lt 10 -and -not $set; $i++) {
  try {
    if ($p.HasExited) { break }
    $p.ProcessorAffinity = [IntPtr]$mask
    $p.PriorityClass     = [System.Diagnostics.ProcessPriorityClass]::$Priority
    $set = $true
  } catch {
    Start-Sleep -Milliseconds 200
    $p.Refresh()
  }
}

if ($set) {
  Write-Host ("[ok] PID {0} lanzado con afinidad {1} core(s) y prioridad {2}." -f $p.Id, $Cores, $Priority) -ForegroundColor Green
  Write-Host "[ok] Espera la linea 'Escuchando conexiones entrantes ...' antes de asumir que arranco." -ForegroundColor Green
} else {
  Write-Host "[AVISO] No se pudo fijar afinidad/prioridad (el proceso salio o nego el handle). Revisa la ventana del server." -ForegroundColor Yellow
}
