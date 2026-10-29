# GridPilot

Piloto de escritorio para Windows x64 y Excel Microsoft 365. Opera un borrador visible mediante operaciones tipadas, texto con Anthropic y dictado Azure Speech en español de México. La exportación exige una acción explícita sobre la revisión completa.

## Ejecutar

La distribución autocontenida se genera en `artifacts/release`. Ejecutar `GridPilot.App.exe`; conservar junto al ejecutable las carpetas `server` y `assets`. Excel de escritorio debe estar instalado. No se necesita instalar .NET para ejecutar la distribución autocontenida.

Para compilar se requiere .NET SDK **10.0.403** (archivo `global.json`). Durante la implementación se instaló localmente en `.tools/dotnet`, sin sustituir SDK del sistema:

```powershell
./scripts/build.ps1 -Publish
./artifacts/release/GridPilot.App.exe
```

En otro equipo, instalar ese SDK desde Microsoft o usar el instalador oficial `dotnet-install.ps1` con `-Version 10.0.403 -InstallDir .tools/dotnet -NoPath`. NuGet requiere Internet durante restauración. Las dependencias directas y transitivas están fijadas en los proyectos y `packages.lock.json`; usar `dotnet restore --locked-mode` en CI.

## Flujo

1. Pulsar **Plantilla demo** o **Abrir libro**. Se toma la versión guardada del archivo, se crea una base de referencia y un borrador, y se abre únicamente este último en una instancia dedicada de Excel.
2. Sin servicios de nube, usar **Lote local** con JSON tipado. Ejemplo:

```json
[
  {"Kind":"SetValue","Table":"Movimientos","Id":"MOV011","Field":"Cantidad","Value":"3"}
]
```

3. Para texto/voz, abrir **Ajustes**, introducir claves en los campos protegidos, región Azure y activar el consentimiento de consumo de nube. El modelo inicial es `claude-sonnet-5-5`; no hay sustitución automática. Las claves se protegen mediante DPAPI del usuario de Windows, fuera del proyecto.
4. Escribir una instrucción, o mantener presionado el botón de voz y soltar. Revisar la transcripción y pulsar **Enviar**. Las instrucciones claras se ejecutan en el borrador; las ambiguas piden aclaración.
5. Revisar cambios y resultados en ambas ventanas. **Actualizar revisión** detecta modificaciones manuales. **Deshacer último lote** restaura el punto previo; no admite deshacer parciales.
6. Pulsar **Aprobar y exportar**, escoger un archivo nuevo y confirmar. Si el borrador cambió después de la revisión, hay que revisarlo otra vez. No se sobrescriben archivos existentes. Sustituir el original es una acción manual fuera de GridPilot.

Los libros que no satisfacen el perfil quedan para inspección. No se afirma compatibilidad de edición con libros arbitrarios. El perfil conocido tiene campos semánticos estables y encabezados editables; los renombrados actualizan el perfil de la sesión.

## Datos y recuperación

La aplicación guarda sesiones en `%LOCALAPPDATA%/GridPilot/sessions`. Cada sesión incluye SQLite, base, borrador, copias válidas y puntos anteriores. **Recuperar sesión** abre su `session.db` y recupera la última copia confirmada; las operaciones pendientes no se repiten. Después de recuperar se exige una nueva aprobación.

**Descartar sesión** cierra el borrador sin exportar y conserva la recuperación local. Se pueden eliminar carpetas de sesiones cerradas manualmente cuando ya no sean necesarias. Los libros y el historial de revisión contienen datos del usuario; son almacenamiento local funcional. Los registros técnicos de consumo omiten celdas, instrucciones, claves y audio. El dictado no conserva archivos de audio locales.

## Contrato de operaciones

`ActionPlan`: `OperationId` GUID, `Intent`, `SourceRevision` y `Operations`. Cada operación tiene `Kind`, `Table`, `Id`, `Field`, `Value` y, para altas, `Values` (diccionario con claves semánticas del perfil). Fechas de escritura `yyyy-MM-dd`, decimales con punto. Las lecturas de fecha contienen seriales numéricos Excel.

Operaciones: `SetValue`, `AddRecord`, `DeleteRecord`, `AddColumn`, `RenameColumn`, `DeleteColumn`, `SetFormula`, `AddSheet`, `RenameSheet`, `DeleteSheet`. `Table` identifica la tabla estable; en operaciones de hoja identifica su nombre. `Value` contiene el nuevo nombre o la fórmula en operaciones estructurales. Nuevas columnas calculadas usan una fórmula A1 relativa a la primera fila de datos (fila 6 en la plantilla); Excel la extiende al resto.

Los IDs existentes no se cambian. Los cálculos esenciales no pueden eliminarse ni sobrescribirse. Las dependencias se validan sobre el lote completo. Nuevas fórmulas admiten referencias internas A1, aritmética y `SUM`, `SUMIFS`, `COUNTIFS`, `IF`, `ROUND`, `INDEX`, `MATCH`; la plantilla usa además `SUMPRODUCT` para contar materiales bajo mínimo.

## Arquitectura

- `GridPilot.Core`: contratos, diferencias y verificador independiente de negocio.
- `GridPilot.Excel`: COM en STA, reintentos acotados de llamadas rechazadas, apertura restringida y snapshots de copias nativas Excel.
- `GridPilot.Storage`: SQLite durable, copias, idempotencia, recuperación, aprobación y exportación.
- `GridPilot.Mcp`: servidor JSON-RPC MCP por stdio, iniciado por la interfaz. Inspección, lecturas acotadas, lote, validación, comparación y restauración. No registra herramientas para aprobar/exportar. Los métodos privados `gridpilot/*` son del controlador local; no se transmiten al modelo.
- `GridPilot.Providers`: Anthropic con herramientas estrictas, límites por instrucción, Azure Speech y DPAPI.
- `GridPilot.App`: WPF, un libro y una instrucción a la vez.

El aislamiento se basa en copias y validación; no es una máquina virtual. No se ofrecen VBA, ejecución arbitraria, Google Sheets, conexiones externas ni colaboración multiusuario.

## Pruebas

```powershell
$env:DOTNET_CLI_HOME="$PWD/.tools/cli"
$env:NUGET_PACKAGES="$PWD/.tools/packages"
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release -- --excel
```

La primera batería usa dobles de prueba; la segunda añade COM real sobre copias desechables y reapertura de la exportación. Los informes se escriben en `artifacts/tests`. Las pruebas de nube se ejecutan por separado según `docs/EVALUATION.md`; no se lanzan al compilar ni al iniciar la aplicación.

Consultar `docs/PROGRESS.md` para resultados, limitaciones y criterios de aceptación. `docs/REFERENCES.md` registra la documentación oficial consultada.
