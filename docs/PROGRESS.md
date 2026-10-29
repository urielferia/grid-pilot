# Estado de GridPilot — 29 de octubre de 2026

La implementación local está disponible y los flujos deterministas se han probado con Excel real. El piloto todavía **no tiene aceptación completa**: faltan las evaluaciones reales de Anthropic y Azure Speech, que requieren credenciales y autorización de consumo. No se realizaron llamadas pagadas.

## Entregas por etapa

| Etapa | Entregado y comprobado | Pendiente |
|---|---|---|
| 1. Estructura y plantilla | Solución C#/.NET 10: WPF, dominio, COM, almacenamiento, MCP, proveedores y pruebas. Plantilla de 4 hojas, 10 materiales, 10 visitas y 20 movimientos. Perfil y valores esperados. Las cuatro hojas se renderizaron y revisaron visualmente; recálculo nativo probado. | Sin bloqueo para la plantilla local. |
| 2. Sesiones y recuperación | Base de referencia, borrador, puntos previos, recuperación explícita, GUID idempotentes, SQLite y confirmación atómica del estado. Cambios manuales invalidan planes y aprobación. | Ensayo manual de terminación forzada exactamente durante una escritura COM; fallos parciales y restauración fallida sí se cubren con dobles de prueba. |
| 3. Ejecución y validación | Operaciones tipadas, registros por ID, datos y estructura, verificador independiente, fórmulas extendidas, rollback y diferencias. MCP inicializa, publica 7 herramientas y cierra por EOF. | Límites de alcance descritos abajo. |
| 4. Conversación y voz | Anthropic con herramientas estrictas y modelo fijo, lecturas acotadas, aclaraciones, límites configurables, DPAPI, Azure Speech es-MX y envío explícito de transcripción. Pruebas simuladas de proveedor desactivado y HTTP 429 sin reintentos ni cambio de modelo. | Cuenta/modelo, audio, red y calidad de interpretación con servicios reales. |
| 5. Revisión y distribución | Revisión agrupada, deshacer, selección de perfil, aprobación completa, copia exacta con SHA-256, exportación a archivo nuevo, binarios autocontenidos Windows x64. Aplicación abre; accesibilidad inspeccionada. | Revisión visual completa de WPF y voz. Captura Windows falló con `FrameArrived timed out` / `window capture timed out`; no se presenta como aprobada. |

## Evidencia ejecutada

- SDK local **10.0.403**, runtime **10.0.14**, Windows x64; Excel **16.0.20326.20232**.
- Compilación Release sin advertencias ni errores; publicación autocontenida completada.
- **34/34 pruebas sin Excel**: dominio, fórmulas, perfiles, proveedores simulados, rollback parcial, restauración fallida bloqueante, transacciones SQLite y persistencia. Informe: `artifacts/tests/20261029-191727/results.json`.
- **59/59 pruebas en la batería con Excel real**, incluidas 32 pruebas de dominio de esa revisión y 27 casos adicionales de integración. Las 2 pruebas nuevas de proveedor se ejecutaron después por separado, sin modificar el ejecutor Excel. Informe: `artifacts/tests/20261029-191522/results.json`.
- Protocolo MCP: `scripts/test-mcp.ps1` comprueba negociación, catálogo sin aprobar/exportar, ping y cierre.
- Prueba de interfaz por accesibilidad/teclado: la distribución abre la plantilla, inicia su servidor e instancia Excel, termina en «Borrador listo» y muestra 8 visitas, 550 MXN y 0 materiales bajo mínimo. La captura gráfica sigue pendiente por fallo del controlador de captura.
- Copias portables de los informes aprobados: `docs/test-results/domain.json` y `docs/test-results/excel.json`.
- Original: SHA-256 conservado en las baterías concluidas.
- Exportación: archivo aprobado reabierto y recalculado en otra instancia de Excel; fórmulas, tablas y reglas verificadas.
- Los informes anteriores que detectaron defectos permanecen en `artifacts/tests` como diagnóstico. No cuentan como resultados aprobados. Se corrigieron el orden de inicialización COM, la consulta de ciclos por hoja, metadatos aleatorios de SaveCopyAs y texto literal que empieza por `=`.

## Criterios de aceptación

| Criterio del plan | Estado |
|---|---|
| Original conserva su huella | Cumple en pruebas ejecutadas. |
| Exportación solo por aprobación explícita | Cumple: sin aprobación se bloquea; ninguna herramienta de modelo aprueba/exporta. Interfaz requiere confirmación humana. |
| Fallos recuperan estado o bloquean sesión | Cumple en rollback real y fallos inyectados; falta ensayo de cierre forzado durante COM. |
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

Configurar claves mediante Ajustes, habilitar nube y seguir `docs/EVALUATION.md`. Probar voz y los 30 casos con `claude-sonnet-5-5`, registrando uso, duración y efectos. Si el modelo no está disponible para la cuenta, se informa el error sin sustituirlo.
