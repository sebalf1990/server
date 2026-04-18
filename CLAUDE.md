# Argentum Online (AO20) — AI Context

## Recursos de referencia
- **Documentación VB6:** https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation
- **Compilador:** `C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe`

---

## Repositorios

| Repositorio | Tipo              | Acceso  | URL |
|-------------|-------------------|---------|-----|
| Assets      | Recursos/Config   | Privado | https://github.com/ao-org/argentum-online-assets/ |
| Server      | Servidor VB6      | Público | https://github.com/ao-org/argentum-online-server/ |
| Client      | Cliente VB6       | Público | https://github.com/ao-org/argentum-online-client  |

- **Assets:** Gráficos, sonidos, mapas, archivos de configuración y recursos compartidos
- **Server:** Lógica del servidor, base de datos, NPCs, sistemas de combate
- **Client:** Cliente del juego, UI, renderizado, manejo de input

---

## Contexto del proyecto

Sos un desarrollador experimentado en Visual Basic 6 trabajando en el MMORPG legacy **Argentum Online**. El proyecto usa arquitectura cliente-servidor, ambos escritos en VB6. La comunicación entre ellos ocurre a través de un protocolo de paquetes personalizado.

**Reglas generales:**
- Seguir siempre prácticas compatibles con VB6 legacy
- Priorizar código limpio y legible con mínimo riesgo de regresión
- **NO usar sintaxis de VB.NET** — solo VB6 puro
- Minimizar cambios en el protocolo cliente-servidor para mantener compatibilidad

---

## Comandos de compilación

### Servidor
```
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d UsarQueSocket=1:ConUpTime=1:AntiExternos=0:Lac=1:DEBUGGING=0:PYMMO=1:UNLOCK_CPU=1:DIRECT_PLAY=0 ^
/out vb6build.log ^
Server.VBP
```

### Cliente
```
"C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe" /make ^
/d Compresion=1:DEBUGGING=1:PYMMO=1:ENABLE_ANTICHEAT=1:REMOTE_CLOSE=1:BATTLESERVER=1:DIRECT_PLAY=0:ENABLE_BASS=1 ^
/out vb6build.log ^
Client.VBP
```

---

## Flujo de trabajo para agentes AI

Cuando modifiques el proyecto, seguí este ciclo:

1. Realizá los cambios en archivos `.bas`, `.cls` o `.frm`
2. Ejecutá el comando de compilación correspondiente (servidor o cliente)
3. Revisá el log `vb6build.log` en busca de errores
4. Si hay errores, analizá el log e intentá corregirlos
5. Repetí hasta que la compilación sea exitosa

---

## Reglas de codificación VB6

### 1. Uso obligatorio de `Call` y paréntesis
Siempre usá `Call` al invocar cualquier Sub. Siempre incluí paréntesis, incluso sin argumentos.

```vb
Call GuardarDatos()
Call EnviarMensaje("Hola", 2)
```

### 2. Funciones siempre con paréntesis
Incluso al ignorar el valor de retorno, las funciones deben llamarse con paréntesis.

```vb
Call ObtenerTiempoActual()
Dim puntos As Long
puntos = CalcularPuntos(usuarioId)
```

### 3. Convenciones de nomenclatura

| Elemento       | Convención         | Ejemplo                   |
|----------------|--------------------|---------------------------|
| Módulos        | Prefijo `mod`      | `modNetwork`, `modLogin`  |
| Formularios    | Prefijo `frm`      | `frmMain`, `frmLogin`     |
| Controles      | Notación húngara   | `txtNombre`, `lblError`   |
| Variables      | camelCase          | `userIndex`, `goldAmount` |
| Constantes     | UPPER_CASE         | `GOLD_PRICE`              |
| Funciones/Subs | PascalCase         | `Call ValidarUsuario()`   |
| Enums          | Prefijo `e_`       | `e_TipoPago`              |

### 4. Espaciado y formato
- Indentación de **4 espacios**
- Una línea en blanco entre procedimientos
- Alinear declaraciones de variables relacionadas:

```vb
Dim userGold    As Long
Dim userSilver  As Long
Dim userName    As String
```

### 5. Sin números mágicos
Declarar todos los valores fijos como constantes en módulos compartidos.

```vb
Public Const GOLD_PRICE As Long = 50000

If .Stats.GLD < GOLD_PRICE Then
    Call EscribirError("No tenés suficiente oro.")
End If
```

### 6. Manejo de errores estandarizado
Siempre usar `On Error GoTo Name_Err` con su label al final del Sub.

```vb
Private Sub ValidarSesion()
    On Error GoTo ValidarSesion_Err

    Call HacerAlgo()

    Exit Sub

ValidarSesion_Err:
    Call TraceError(Err.Number, Err.Description, "modLogin.ValidarSesion", Erl)
End Sub
```

### 7. Identificadores explícitos
Usar nombres claros y específicos. Evitar nombres genéricos como `dato`, `res`, `temp`.

```vb
Dim creditAmount As Long
Dim connectionId As Integer
```

### 8. Queries SQL
Siempre usar queries parametrizadas (`?`). Siempre cerrar todos los objetos `Recordset`.

```vb
Dim RS As ADODB.Recordset
Set RS = Query("SELECT nivel FROM user WHERE id = ?;", userId)

If Not RS.EOF Then
    nivel = RS!nivel
End If

Call RS.Close
```

### 9. Mensajes localizados

#### Servidor (`WriteLocaleMsg`)
El servidor envía un ID de mensaje; el cliente lo resuelve desde un archivo de índice.

```vb
Call WriteLocaleMsg(UserIndex, "1291", FONTTYPE_INFOBOLD, GOLD_PRICE)
```
`"1291"` mapea a: *"Necesitás al menos ¬1 de oro"* — `GOLD_PRICE` reemplaza `¬1`.

Archivos de mensajes:
- Español: `SP_LocalMsg.dat`
- Inglés: `EN_LocalMsg.dat`

Para regenerar: ejecutar `python generar_localindex.py` o `generar_localindex.exe` desde `/tools/`.

#### Cliente (`JsonLanguage`)
Traducción de UI en archivos JSON dentro de la carpeta `Languages/`:
- `Languages/1.json` (Español)
- `Languages/2.json` (Inglés)

```vb
Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_CARGAR_OPCIONES"), vbCritical, JsonLanguage.Item("TITULO_ERROR_CARGAR"))
```

### 10. Control de flujo claro
Usar `Exit Sub` para salidas tempranas. Evitar anidamiento innecesario.

```vb
If Not EstaAutenticado(UserIndex) Then
    Call Desconectar(UserIndex)
    Exit Sub
End If
```

### 11. Funciones puras y sin efectos secundarios
Preferir `Function` para validaciones o transformaciones. Evitar modificar variables globales innecesariamente.

---

## Trazabilidad y control de versiones

- Usar **Git** para registrar todos los cambios
- Antes de cada feature nueva, crear una rama: `git checkout -b feature/nombre-feature`
- Hacer commit descriptivo antes y después de cada cambio significativo
- Formato de commit recomendado: `[SERVER|CLIENT] Descripción del cambio`

---

## Tips para modificaciones

- Antes de modificar, identificá el módulo relevante explorando la estructura del proyecto
- Seguí siempre el patrón existente en el código — no introducir nuevos patrones sin consenso
- Documentar cada cambio con un comentario en el código indicando qué se modificó y por qué
- Testear en servidor local antes de cualquier merge
