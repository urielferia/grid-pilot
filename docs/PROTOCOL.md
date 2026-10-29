# Transporte local

GridPilot inicia `GridPilot.Mcp.exe` como proceso hijo, con stdin/stdout redirigidos. Se implementan mensajes JSON-RPC 2.0 delimitados por nueva línea, negociación MCP `2025-06-18`, `initialize`, `notifications/initialized`, `ping`, `tools/list` y `tools/call`. No hay puerto TCP, servidor remoto ni endpoint de red local.

Herramientas del ejecutor: `inspect_workbook`, `read_records`, `read_range`, `execute_batch`, `recalculate_validate`, `compare_revision` y `restore_checkpoint`. Esta última restaura exclusivamente el último lote; no acepta rutas arbitrarias. Lecturas limitadas a 20 registros o 100 celdas. Lotes limitados a 100 operaciones. Libro limitado a 2000 filas y 50 columnas por hoja.

El controlador local usa además `gridpilot/open`, `gridpilot/recover`, `gridpilot/approve`, `gridpilot/export` y `gridpilot/close`. Estos métodos no aparecen en el catálogo MCP ni en las herramientas Anthropic. No se debe exponer el proceso a clientes no confiables: el límite de confianza es el controlador local de este usuario de Windows, no la posesión de un nombre de método secreto.

Anthropic recibe el perfil y la instrucción, y puede solicitar lecturas acotadas de registros. Las celdas se encapsulan como resultados de herramientas de datos. El único resultado operativo del modelo es `propose_plan`: el controlador fija su GUID y revisión a partir del contexto local. Ningún texto del modelo se interpreta como código ni como una ruta de exportación.

Todas las llamadas COM pertenecen al mismo hilo STA. Un filtro OLE reintenta únicamente llamadas rechazadas por Excel ocupado durante cinco segundos; no se reintenta un lote parcialmente ejecutado. Durante la ejecución/aprobación se deshabilita la interacción en esa instancia dedicada. La revisión usa valores, fórmulas expandidas por Excel y XML nativo (estructura, estilos y propiedades), excluyendo selección, posición de ventana y metadatos de identidad de SaveCopyAs.

SQLite registra el intento antes de escribir y conserva la huella del plan. Reutilizar el ID con contenido distinto falla. Los fallos restauran la copia previa; si no se puede restaurar, la sesión se bloquea. La recuperación explícita usa la última copia confirmada y no repite intentos pendientes.
