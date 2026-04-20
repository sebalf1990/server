# Cambios realizados sobre el código original AO20

## Resumen

Este fork parte del commit `796e1e3` ("código original AO20 sin modificar") y corrige
cuatro bugs que impedían que el servidor arrancara en una instalación limpia de Windows 10.

---

## Bugs corregidos

### 1. `Subscript out of range` en `LoadGlobalQuests` — `ModGlobalQuests.bas`

**Problema:** Al acceder a `RS!total_amount` sobre un Recordset vacío (tabla
`global_quest_user_contribution` sin filas) se producía error 9 en VB6.

**Causa:** La query `SELECT SUM(amount) AS total_amount ... WHERE event_id = ?` con
cursor forward-only devuelve un RS posicionado al EOF cuando no hay filas, y `RecordCount`
retorna -1. El código accedía al campo sin validar.

**Fix en `Codigo/ModGlobalQuests.bas` (~línea 126):**
```vb
' Antes
If Not IsNull(RS!total_amount) Then
    .GatheringGlobalCounter = RS!total_amount
    .GatheringGlobalInstallments = RS!total_amount + .GatheringInitialInstallments
End If

' Después
If Not RS Is Nothing Then
    If RS.RecordCount > 0 Then
        If Not IsNull(RS!total_amount) Then
            .GatheringGlobalCounter = RS!total_amount
            .GatheringGlobalInstallments = RS!total_amount + .GatheringInitialInstallments
        End If
    End If
    Call RS.Close
    Set RS = Nothing
End If
```

También se corrigió la clave INI de `BossIndex` → `NPCID` para que coincida con
`GlobalQuests.dat`.

---

### 2. `LoadPhoenixModule` cuelga el servidor — `ModPhoenix.bas` + `PhoenixMapPool.dat`

**Problema:** El servidor se colgaba permanentemente al llegar a `LoadPhoenixModule`.

**Causa:** El archivo `../Recursos/Dat/PhoenixMapPool.dat` tenía line endings Unix (LF).
El parser INI de VB6 (`clsIniManager`) usa `Line Input #handle` que en Windows **solo
trata CR o CRLF como fin de línea** — no LF solo. Resultado: todo el archivo se leía
como una línea gigante, ninguna sección era parseada, `GetValue("INIT", "MaxPhoenixMaps")`
devolvía `""`, `val("")` = 0, y `ReDim Preserve PhoenixMapPool(1 To 0)` lanzaba error 9.
Sin handler en el Sub, el error llegaba a `Sub Main`, que al salir dejaba las ventanas
abiertas sin game loop = servidor "colgado".

**Fix 1 — `../Recursos/Dat/PhoenixMapPool.dat`:** Convertir a CRLF.

**Fix 2 — `Codigo/ModPhoenix.bas`:** Agregar `On Error GoTo` + guard defensivo:
```vb
Public Sub LoadPhoenixModule()
    On Error GoTo LoadPhoenixModule_Err
    ...
    If MaxPhoenixMaps < 1 Then
        Call LogError("PhoenixMapPool.dat: MaxPhoenixMaps invalido: " & MaxPhoenixMaps)
        Exit Sub
    End If
    ...
LoadPhoenixModule_Err:
    Call TraceError(Err.Number, Err.Description, "ModPhoenix.LoadPhoenixModule", Erl)
End Sub
```

---

### 3. Conexión ADO falla con error 3706 al lanzar desde algunos entornos

**Problema:** `Connection.Open` con `DRIVER={SQLite3 ODBC Driver}` fallaba con error 3706
("Provider cannot be found") cuando el servidor se lanzaba desde el Explorador de Windows
o desde VS Code.

**Causa:** El proveedor MSDASQL (OLE DB para ODBC) no se cargaba correctamente desde ese
contexto de ejecución. Los componentes OLE DB del sistema necesitaban re-registro.

**Fix — acción en el sistema (no en código):** Re-registrar la DLL de OLE DB:
```
regsvr32 "C:\Program Files (x86)\Common Files\System\Ole DB\oledb32.dll"
```
> **Nota:** El servidor **debe lanzarse siempre desde un cmd** (ver instrucciones abajo).
> Desde el Explorador o VS Code el contexto de carga de DLLs difiere y puede reproducir
> el error en algunos sistemas.

---

### 4. `modDatabase.bas` — mejoras menores

- Se removió código de diagnóstico temporal.
- Sin cambios funcionales al flujo de conexión.

---

## Archivos modificados

| Archivo | Tipo de cambio |
|---|---|
| `Codigo/ModGlobalQuests.bas` | Bug fix — guard RS vacío en `LoadGlobalQuests` |
| `Codigo/ModPhoenix.bas` | Bug fix — error handler + guard `MaxPhoenixMaps` |
| `Codigo/modDatabase.bas` | Limpieza — sin cambio funcional |
| `Codigo/General.bas` | Limpieza — sin cambio funcional |
| `../Recursos/Dat/PhoenixMapPool.dat` | Fix encoding — LF → CRLF |

---

# Instructivo de arranque del servidor

## Requisitos previos

| Componente | Versión | Notas |
|---|---|---|
| Windows 10 (32 o 64 bit) | — | |
| Visual Basic 6.0 IDE | VB98 | Solo para compilar |
| SQLite3 ODBC Driver (32-bit) | — | Instalar desde http://www.ch-werner.de/sqliteodbc/ → `sqliteodbc.exe` |
| Microsoft MDAC / OLE DB 32-bit | — | Incluido en Windows, re-registrar si hay error 3706 |

### Re-registro de OLE DB (hacer una sola vez si hay error 3706)

Abrir cmd como Administrador y ejecutar:

```
regsvr32 "C:\Program Files (x86)\Common Files\System\Ole DB\oledb32.dll"
regsvr32 "C:\Program Files (x86)\Common Files\System\ado\msado15.dll"
```

---

## Estructura de carpetas esperada

```
C:\AO20\
  argentum-online-server-5.6.35\   ← este repositorio
  Recursos\
    Dat\
    Mapas\
    ...
```

---

## Setup inicial (primera vez)

1. Clonar el repositorio en `C:\AO20\argentum-online-server-5.6.35\`

2. Copiar archivos de configuración de ejemplo:
   ```
   copy Example.Server.ini Server.ini
   copy Example.feature_toggle.ini feature_toggle.ini
   copy Example.Configuracion.ini Configuracion.ini
   ```

3. Copiar la DB vacía como DB de trabajo:
   ```
   copy Empty_db.db Database.db
   ```

4. Copiar el motd de ejemplo:
   ```
   copy ..\Recursos\Dat\Example.Motd.ini ..\Recursos\Dat\Motd.ini
   ```

5. Compilar el servidor (ver sección siguiente).

---

## Compilar el servidor

Desde un **cmd normal** (no hace falta Admin), dentro de la carpeta del server:

```bat
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d "UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=1:UNLOCK_CPU=1:DIRECT_PLAY=0" ^
/out vb6build.log ^
Server.VBP
```

O usar el atajo `_build_tmp.bat` incluido en el repo.

Revisar `vb6build.log` — debe decir `Build of 'server.exe' succeeded.`

---

## Arrancar el servidor

> **IMPORTANTE:** El servidor **siempre debe lanzarse desde un cmd**, no desde el
> Explorador de Windows ni desde VS Code. Esto es necesario para que el driver ODBC de
> SQLite se cargue correctamente.

```bat
cd C:\AO20\argentum-online-server-5.6.35
server.exe
```

El servidor arranca con la ventana `frmCargando` y luego muestra:
```
Escuchando conexiones entrantes ...
```

Las migraciones SQL de `ScriptsDB/` se aplican automáticamente al primer arranque.

---

## Flags de compilación disponibles

| Flag | Valor por defecto | Descripción |
|---|---|---|
| `UsarQueSocket` | 1 | Usar socket queue (recomendado) |
| `ConUpTime` | 1 | Mostrar uptime |
| `AntiExternos` | 0 | Anti-cheat externo (desactivado) |
| `Lac` | 1 | LAC habilitado |
| `DEBUGGING` | 0 | Debug mode (activa Debug.Print) |
| `PYMMO` | 1 | Features PYMMO |
| `UNLOCK_CPU` | 1 | Sin limitación de CPU |
| `DIRECT_PLAY` | 0 | DirectPlay desactivado (usar sockets) |

---

## Troubleshooting

### Error 3706 "Provider cannot be found"
El servidor debe lanzarse **desde cmd**. Si el error persiste, re-registrar OLE DB:
```
regsvr32 "C:\Program Files (x86)\Common Files\System\Ole DB\oledb32.dll"
```

### Servidor se cuelga en "Cargando Quests"
Verificar que todos los archivos `.dat` en `../Recursos/Dat/` tengan line endings CRLF
(Windows). Archivos con LF solo (Unix) no son parseados correctamente por el INI reader
de VB6.

### Error 404 "Mapa corrupto o no existe"
Verificar que todos los mapas en `../Recursos/Mapas/` tengan numeración consecutiva sin
huecos (ej: si existen Mapa1..Mapa842 y Mapa844, falta Mapa843).

### Error 9 "Subíndice fuera del intervalo" al arrancar con DB nueva
Las tablas de quests globales se crean via migraciones automáticas. Si `Database.db` fue
copiado de `Empty_db.db`, las migraciones corren solas al primer arranque.
