# Estado de GridPilot — actualización del 30 de octubre de 2026

La implementación local está disponible y los flujos deterministas se han probado con Excel real. El piloto todavía **no tiene aceptación completa**: faltan las evaluaciones reales de Gemini/Anthropic y Azure Speech, que requieren credenciales y autorización de consumo. No se realizaron llamadas pagadas.

## Trabajo de esta revisión

- Se comparó el plan con los módulos y la evidencia existentes. Etapas 1–3 están implementadas y verificadas; etapas 4–5 tienen todos sus flujos locales, de recuperación y de distribución completos y probados, quedando pendiente únicamente la aceptación con servicios reales de nube/voz.
- Se diagnosticó y corrigió un defecto de codificación UTF-8 en las tuberías stdio entre `RpcClient` (proceso GUI sin consola, que por defecto usa la página de códigos OEM/ANSI en Windows) y `GridPilot.Mcp`, que corrompía los encabezados acentuados del perfil (`Técnico`, `Descripción`, `Mínimo`) y marcaba la plantilla como solo lectura (`Editable: false`) cuando la página de códigos no era UTF-8. Ahora tanto `RpcClient` como `GridPilot.Mcp` y `scripts/test-mcp.ps1` fuerzan UTF-8 sin BOM.
- Se implementó y probó la recuperación limpia ante terminación forzada del servidor MCP o de Excel durante una escritura COM (`Owner` en `src/GridPilot.Mcp/Owner.cs`, seguimiento de PID en `ExcelWorkbook`, reinicio de `RpcClient` en `MainWindow` y liberación de `Journal` en constructores de `Session`). Si el servidor termina inesperadamente durante un lote y deja huérfana su instancia de Excel bloqueando `draft.xlsx`, `gridpilot/recover` cierra únicamente esa instancia huérfana creada por GridPilot y restaura el último punto válido sin repetir escrituras pendientes.
- Se endureció `scripts/build.ps1` y `Directory.Build.props` (`NuGetAudit=false`, `APPDATA`/`LOCALAPPDATA` locales bajo `.tools/cli` y `-nodeReuse:false -m:1`) para que la compilación y publicación funcionen de forma reproducible sin red ni acceso fuera del repositorio.
- Se añadió Gemini Developer API mediante HTTP nativo y herramientas tipadas. Usa el mismo contrato de planificación y ejecutor local que Anthropic. Solo expone lecturas acotadas y propuesta de plan; nunca aprobación/exportación.
- Ajustes permite elegir proveedor/modelo y guardar la clave Gemini con DPAPI. Instalaciones nuevas sugieren gemini-3.8-flash; la configuración antigua conserva Anthropic. No se modifica ninguna clave ni se activa nube automáticamente.
- Se preservan firmas opacas/IDs de llamadas Gemini entre turnos. Se contabilizan tokens de pensamiento, se rechazan respuestas truncadas/bloqueadas y se mantienen límites por instrucción sin reintentos/fallback pagados.
- Se reforzó el parser compartido: operaciones conocidas, máximo de 100 y aclaraciones sin escrituras.
- El SDK no estaba en este checkout. Se instaló el SDK fijado 10.0.403 localmente y se restauraron las dependencias existentes; no se añadieron paquetes.
- Compilación Release sin advertencias ni errores. **54/54 pruebas locales** y **84/84 en la batería con Excel real** aprobadas (incluyendo cierre inesperado de Excel + recuperación, terminación forzada del servidor MCP durante escritura COM + recuperación sin reintento, apertura de plantilla editable por tubería RPC UTF-8 e instrucciones complejas de reemplazo/intercambio/eliminación en cascada de filas); MCP pasa negociación, catálogo de 7 herramientas, ping y EOF. Los tests de proveedor usan HTTP simulado, sin consumo.
- Informes nuevos: `artifacts/tests/20261031-061659/results.json` y `artifacts/tests/20261031-061709/results.json`; copias versionadas en `docs/test-results/domain.json` y `docs/test-results/excel.json`. Original conserva SHA-256 (`BCEC82F38F66EA73EDE98D3E7D41BC00BE16892709547493D0C07A71D5E33D51`) y exportación reabierta/recalculada en Excel.
- Se generó distribución autocontenida Windows x64 en artifacts/release. La comprobación visual detectó ventanas que excedían la altura de pantalla; se ajustó su altura al área de trabajo y se agregó desplazamiento al lote local.
- Smoke test nuevo: scripts/test-mcp.ps1 -Excel -Published aprobó apertura de copia demo, perfil editable, recálculo/validación y cierre con el servidor autocontenido; SHA-256 original intacto. Esto verifica el servidor distribuido además de la biblioteca probada por la batería.
- La verificación por tubería RPC UTF-8 (`UiRpcDemoOpen`) y `scripts/test-mcp.ps1 -Excel -Published` confirmaron la apertura editable de la plantilla demo desde el servidor publicado; los errores COM siguen informando HRESULT sin exponer datos de celdas ni secretos.
- Se documentó asignación de cada herramienta/modelo, costo ilustrativo con USD 5 y papel opcional de JEV en docs/MODELS.md. El usuario confirmó crédito de AI Studio. JEV no se integra como dependencia obligatoria porque el piloto no necesita otro clasificador.

## Entregas por etapa (base original y pendientes)

| Etapa | Entregado y comprobado | Pendiente |
|---|---|---|
| 1. Estructura y plantilla | Solución C#/.NET 10: WPF, dominio, COM, almacenamiento, MCP, proveedores y pruebas. Plantilla de 4 hojas, 10 materiales, 10 visitas y 20 movimientos. Perfil y valores esperados. Las cuatro hojas se renderizaron y revisaron visualmente; recálculo nativo probado. | Sin bloqueo para la plantilla local. |
| 2. Sesiones y recuperación | Base de referencia, borrador, puntos previos, recuperación explícita, GUID idempotentes, SQLite, confirmación atómica del estado, cierre de instancia Excel huérfana creada por la sesión y ensayo automático de terminación forzada durante escritura COM sin repetir escrituras. Cambios manuales invalidan planes y aprobación. | Sin bloqueo local. |
| 3. Ejecución y validación | Operaciones tipadas, registros por ID, datos y estructura, verificador independiente, fórmulas extendidas, rollback y diferencias. MCP inicializa, publica 7 herramientas y cierra por EOF. | Límites de alcance descritos abajo. |
| 4. Conversación y voz | Gemini/Anthropic con planes estructurados y modelo fijo, lecturas acotadas, aclaraciones, límites configurables, DPAPI, Azure Speech es-MX y envío explícito de transcripción. Pruebas simuladas de proveedor desactivado y HTTP 429 sin reintentos ni cambio de modelo. | Cuenta/modelo, audio, red y calidad de interpretación con servicios reales. |
| 5. Revisión y distribución | Revisión agrupada, deshacer, selección de perfil, aprobación completa, copia exacta con SHA-256, exportación a archivo nuevo, binarios autocontenidos Windows x64. Aplicación abre; accesibilidad inspeccionada y apertura demo por tubería RPC UTF-8 verificada. | Aceptación de voz y batería de 30 casos en nube según `docs/EVALUATION.md`. |

## Evidencia histórica de la revisión anterior (29/10/2026)

- SDK local **10.0.403**, runtime **10.0.14**, Windows x64; Excel **16.0.20326.20232**.
- Compilación Release sin advertencias ni errores; publicación autocontenida completada.
- **34/34 pruebas sin Excel**: dominio, fórmulas, perfiles, proveedores simulados, rollback parcial, restauración fallida bloqueante, transacciones SQLite y persistencia. Informe: `artifacts/tests/20261029-191727/results.json`.
- **59/59 pruebas en la batería con Excel real**, incluidas 32 pruebas de dominio de esa revisión y 27 casos adicionales de integración. Las 2 pruebas nuevas de proveedor se ejecutaron después por separado, sin modificar el ejecutor Excel. Informe: `artifacts/tests/20261029-191522/results.json`.
- Protocolo MCP: `scripts/test-mcp.ps1` comprueba negociación, catálogo sin aprobar/exportar, ping y cierre.
- Prueba de interfaz por accesibilidad/teclado: la distribución abre la plantilla, inicia su servidor e instancia Excel, termina en «Borrador listo» y muestra 8 visitas, 550 MXN y 0 materiales bajo mínimo. En esa revisión la captura gráfica falló; en la revisión del 30/10 volvió a funcionar.
- Copias portables actualizadas con la última revisión: `docs/test-results/domain.json` y `docs/test-results/excel.json`.
- Original: SHA-256 conservado en las baterías concluidas.
- Exportación: archivo aprobado reabierto y recalculado en otra instancia de Excel; fórmulas, tablas y reglas verificadas.
- Los informes anteriores que detectaron defectos permanecen en `artifacts/tests` como diagnóstico. No cuentan como resultados aprobados. Se corrigieron el orden de inicialización COM, la consulta de ciclos por hoja, metadatos aleatorios de SaveCopyAs y texto literal que empieza por `=`.

## Criterios de aceptación

| Criterio del plan | Estado |
|---|---|
| Original conserva su huella | Cumple en pruebas ejecutadas. |
| Exportación solo por aprobación explícita | Cumple: sin aprobación se bloquea; ninguna herramienta de modelo aprueba/exporta. Interfaz requiere confirmación humana. |
| Fallos recuperan estado o bloquean sesión | Cumple: probado en rollback real, fallos inyectados, cierre inesperado de Excel y terminación forzada del servidor MCP durante escritura COM con recuperación explícita sin reintento. |
| Cálculos independientes y tolerancias | Cumple: cantidades `0.000001`, importes `0.01 MXN`, IDs/conteos exactos. |
| 30 instrucciones, ≥90 % claras, todas ambiguas/inválidas | Batería fija preparada en `tests/evaluation-30.json`; evaluación real pendiente. No se informa precisión no medida. |
| Cuatro hojas revisadas y exportación reabierta | Cumple: previews de plantilla y comprobación nativa de exportación. |
| Simulaciones y servicios reales separados | Cumple la separación; nube real pendiente. Uso y tiempos se registrarán sin celdas ni credenciales. |

## Límites y supuestos

- Un usuario, un libro, una instrucción. `.xlsx` local con las tres tablas semánticas del proceso; otros libros solo se editan si el perfil compatible pasa la validación. Resumen conserva controles B5/B6 y resultados B9:B11.
- Máximo 2000 filas y 50 columnas por hoja, 100 operaciones por lote, 20 registros por búsqueda y 100 celdas por lectura. Textos enviados al modelo se acotan a 500 caracteres por celda con aviso de truncamiento.
- Se conserva al menos una fila por tabla para garantizar la recuperación de fórmulas. Eliminar el último registro se rechaza explícitamente; todavía no se admiten tablas completamente vacías.
- Nuevos cálculos en columnas adicionales. Fórmulas esenciales protegidas. Fórmulas nuevas: A1 internas, aritmética y funciones permitidas; la plantilla usa además SUMPRODUCT.
- Comparación de valores, fórmulas, estructura, estilos y XML de propiedades; excluye selección/ventana e identidad de copia. No admite macros, Power Query, vínculos externos, pivots, ActiveX ni objetos incrustados.
- Excel instalado es requisito. Distribución portable, sin firma de instalador; conservar `server` y `assets` junto al ejecutable.
- Límites de tokens por instrucción con uso observado y máximo de salida; no equivalen a un límite monetario de cuenta. Configurar también cuotas del proveedor.
- Copias y auditoría funcional local para recuperación; no se conserva audio local. No es una máquina virtual de seguridad.

## Validación externa pendiente

Configurar claves mediante Ajustes, habilitar nube y seguir `docs/EVALUATION.md`. Probar voz y los 30 casos con el modelo seleccionado (inicialmente Gemini `gemini-3.8-flash`), registrando uso, duración y efectos. Si el modelo no está disponible para la cuenta, se informa el error sin sustituirlo.
