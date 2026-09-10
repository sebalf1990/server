@echo off
rem restart_server.bat - relanzador del server (plan 09.001, Ola 1)
rem Uso: restart_server.bat <PID_viejo> <modo>   modo = normal o after-build
rem Lo lanza modSafeRestart.ReinicioSeguro via Shell(..., vbHide) DESDE el
rem propio server.exe elevado: este proceso hijo hereda esa elevacion, asi
rem que el server.exe que arranca al final NO pide UAC de nuevo.
rem Espera a que el PID viejo desaparezca (tope 60 x ~1s) para no chocar con
rem el MsgBox de doble instancia de Main() (GetProcessCount). En modo
rem after-build espera ademas build.done (tope 900 x ~1s, ~15 min) antes de
rem lanzar, para el ciclo compilar->relanzar sin intervencion manual.
rem No usa timeout: falla sin consola interactiva heredada de un Shell sin
rem ventana (decision Ola 0, ver bitacora). Se usa ping -n 2 127.0.0.1 en su
rem lugar. Cada paso se loguea en Logs\safe-restart.log.

set PID=%1
set MODO=%2
if "%MODO%"=="" set MODO=normal
set DIR=%~dp0
set LOGFILE=%DIR%Logs\safe-restart.log

echo %date% %time%;bat_inicio;;%MODO%;%PID%>>"%LOGFILE%"

set /a INTENTOS=0

:esperar_pid
if %INTENTOS% GEQ 60 goto timeout_pid
tasklist /FI "PID eq %PID%" 2>nul | find "%PID%" >nul
if errorlevel 1 goto pid_muerto
ping -n 2 127.0.0.1 >nul
set /a INTENTOS+=1
goto esperar_pid

:timeout_pid
echo %date% %time%;bat_timeout_pid;;%MODO%;%PID%>>"%LOGFILE%"
exit /b 1

:pid_muerto
if /I "%MODO%"=="after-build" goto esperar_build_setup
goto lanzar

:esperar_build_setup
set /a INTENTOS_BUILD=0

:esperar_build_loop
if %INTENTOS_BUILD% GEQ 900 goto timeout_build
if exist "%DIR%build.done" goto build_listo
ping -n 2 127.0.0.1 >nul
set /a INTENTOS_BUILD+=1
goto esperar_build_loop

:timeout_build
echo %date% %time%;bat_timeout_build;;%MODO%;%PID%>>"%LOGFILE%"
exit /b 2

:build_listo
del "%DIR%build.done"

:lanzar
if not exist "%DIR%server.exe" (
    echo %date% %time%;bat_no_exe;;%MODO%;%PID%>>"%LOGFILE%"
    exit /b 3
)

start "" /D "%DIR%" "%DIR%server.exe"
echo %date% %time%;bat_lanzado;;%MODO%;%PID%>>"%LOGFILE%"
exit /b 0
