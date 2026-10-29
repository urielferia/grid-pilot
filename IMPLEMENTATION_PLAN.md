# GridPilot: plan de implementación

## Instrucciones para Codex

Implementa este proyecto siguiendo las decisiones y etapas de este documento. Trabaja en el repositorio actual y respeta cualquier instrucción aplicable de `AGENTS.md`.

- Inspecciona el estado del proyecto antes de modificarlo y conserva el trabajo existente.
- Avanza por etapas en orden. Completa y verifica cada entrega antes de continuar con las capacidades que dependen de ella.
- Mantén una lista de progreso con tareas completadas, pruebas ejecutadas y bloqueos reales. No declares completada una integración que solo funciona con simulaciones.
- No pidas confirmar decisiones ya establecidas aquí. Resuelve detalles menores de implementación con criterio y documenta los supuestos relevantes.
- No inicies compras, suscripciones ni consumo de servicios pagados sin credenciales y autorización de uso. Nunca solicites pegar secretos en el chat ni los incluyas en archivos versionados.
- Si faltan credenciales o dependencias externas, implementa los adaptadores y usa proveedores simulados para continuar las pruebas independientes. Documenta los pasos pendientes para habilitar la integración real.
- No alteres archivos originales del usuario. Ejecuta pruebas de Excel únicamente sobre copias desechables.
- No amplíes el alcance a otras plataformas, macros o servicios compartidos sin una nueva instrucción.
- Consulta documentación oficial vigente antes de integrar APIs y fija versiones de dependencias. No sustituyas silenciosamente el modelo acordado si no está disponible.
- Al finalizar, entrega instrucciones de instalación y uso, resultados de pruebas, limitaciones y estado de cada criterio de aceptación.

Este documento autoriza la implementación del proyecto; la aprobación humana descrita a continuación pertenece al funcionamiento del producto y no implica pedir permiso para cada edición del código.

## 1. Objetivo y alcance de la primera versión

Construir una aplicación para Windows que permita operar Excel mediante voz y texto. El agente analizará el libro, interpretará instrucciones y ejecutará cambios sobre una copia visible en Excel. El usuario revisará el resultado y aprobará la exportación de un archivo nuevo.

### Decisiones confirmadas

- Excel de escritorio en Windows, con una ventana de GridPilot junto a Excel.
- Plantilla inicial de visitas de servicio técnico e inventario.
- Un almacén, sin existencias negativas.
- Edición de datos, registros, fórmulas, columnas y hojas.
- Servicios de IA en la nube, enviando únicamente el contexto necesario.
- Revisión y aprobación por el mismo usuario.
- Exportación de una copia aprobada; sustitución manual del original.

### Límites iniciales

Un usuario, un libro y una instrucción en ejecución a la vez. Archivos `.xlsx` locales. Quedan fuera Google Sheets, macros, Power Query, conexiones externas, colaboración simultánea y administración de usuarios.

La aplicación empezará con una plantilla conocida y un perfil de reglas. Otros libros podrán inspeccionarse, pero solo podrán editarse después de configurar y validar su perfil de compatibilidad.

## 2. Arquitectura e interfaces

| Componente | Implementación |
| --- | --- |
| Aplicación | C# y WPF sobre .NET 10 LTS, distribuida para Windows x64. |
| Control de Excel | Proceso local mediante COM, con una instancia de Excel dedicada exclusivamente al borrador. |
| Herramientas del agente | Servidor MCP local mediante comunicación estándar entre procesos, iniciado por GridPilot. |
| Interpretación | API de Anthropic, inicialmente `claude-sonnet-5-5`, con herramientas y respuestas estructuradas. |
| Voz | Azure Speech, español de México, con botón para mantener presionado y hablar. |
| Persistencia | SQLite para sesiones y operaciones; archivos locales para originales de referencia y puntos de recuperación. |
| Credenciales | Protección por usuario mediante Windows DPAPI; sin claves en el repositorio ni en registros. |

.NET 10 tiene soporte LTS. Durante la planificación se confirmó Excel Microsoft 365 x64 en el equipo y únicamente SDK de .NET anteriores a la versión requerida. Verificar de nuevo las dependencias al comenzar e instalar o documentar la instalación del SDK de .NET 10 según los permisos disponibles.

COM permite conservar el cálculo nativo de Excel y guardar copias. Todas sus operaciones se ejecutarán de forma secuencial en un hilo STA, con manejo de Excel ocupado y liberación de recursos. La aplicación solo cerrará la instancia de Excel que ella haya creado.

El modelo inicial será configurable y se fijará explícitamente; cambiarlo requerirá repetir las evaluaciones del piloto. Las respuestas estructuradas validan el formato, mientras que GridPilot verificará su significado y sus efectos.

### Contratos internos principales

- `WorkbookProfile`: tablas, campos, identificadores, relaciones, zonas editables, fórmulas admitidas y reglas del proceso.
- `WorkbookSnapshot`: revisión del libro, estructura, valores, fórmulas y propiedades relevantes para comparar cambios.
- `ActionPlan`: intención interpretada, revisión de origen y operaciones propuestas.
- `ChangeSet`: cambios directos, efectos calculados, validaciones y punto de recuperación.
- `ApprovalRecord`: revisión aprobada, fecha y huella del archivo exportado.

### Herramientas MCP

- Inspeccionar libro, leer rangos y buscar registros.
- Ejecutar un lote de operaciones tipadas.
- Recalcular, validar, comparar revisiones y restaurar un punto anterior.

Las operaciones incluirán escritura de valores y fórmulas, altas y bajas de registros y creación, eliminación o renombrado de columnas y hojas. No se ofrecerá ejecución arbitraria de código, comandos del sistema o VBA.

La aprobación y la exportación estarán en el controlador de la interfaz, fuera de las herramientas disponibles para el modelo. Las skills describirán procedimientos; las restricciones se aplicarán en código.

## 3. Plantilla de visitas e inventario

Crear un archivo de referencia con cuatro hojas y datos ficticios identificados como demostración:

| Hoja | Contenido |
| --- | --- |
| Resumen | Fechas del reporte, visitas realizadas, costo de materiales del período y materiales actualmente bajo su mínimo. |
| Visitas | ID, fecha, cliente, sitio, técnico, estado, descripción del trabajo y costo calculado de materiales. |
| Inventario | ID del material, descripción, unidad, saldo inicial, mínimo, costo de referencia, entradas, consumos y existencia calculada. |
| Movimientos | ID, fecha, tipo, material, cantidad, visita relacionada, costo unitario e importe calculado. |

La plantilla incluirá 10 materiales, 10 visitas y 20 movimientos consistentes. Usará tablas de Excel, filtros, encabezados inmovilizados, formatos de fecha `dd/mm/aaaa`, importes en MXN y distinción visual entre entradas y cálculos.

### Reglas del piloto

- Existencia = saldo inicial + entradas − consumos.
- Cantidades positivas; identificadores únicos y estables.
- Los consumos requieren una visita existente con estado «Realizada». Las entradas no requieren visita. Los tipos de movimiento serán «Entrada» y «Consumo»; los estados iniciales de visita serán «Programada» y «Realizada».
- El costo unitario se guarda en cada movimiento; inicialmente se propone el costo de referencia del material. Los cambios posteriores al catálogo no alteran costos históricos.
- El costo de una visita suma sus consumos; los movimientos conservan su fecha propia.
- Se valida el saldo resultante del lote completo. La primera versión no reconstruye disponibilidad histórica por fecha.
- Eliminar visitas o materiales con movimientos asociados se bloquea, salvo que la instrucción incluya explícitamente resolver esas dependencias en el mismo lote.
- No se permiten cambios que eliminen identificadores, relaciones o cálculos indispensables para el proceso.

Las modificaciones estructurales podrán agregar campos, reorganizar hojas y ampliar cálculos. Renombrar campos actualizará sus referencias y el perfil. Las fórmulas nuevas se limitarán inicialmente a referencias internas, operaciones aritméticas y funciones comunes como `SUM`, `SUMIFS`, `COUNTIFS`, `IF`, `ROUND`, `INDEX` y `MATCH`. Las funciones externas, volátiles o no soportadas producirán una explicación de incompatibilidad.

La plantilla se generará como artefacto `.xlsx`, se revisará visualmente y se probará con recálculo en Excel. Las fechas y cantidades se almacenarán como valores tipados; los identificadores serán texto. Las fórmulas deberán extenderse correctamente al agregar registros.

## 4. Implementación por etapas

### Etapa 1: estructura del proyecto y plantilla

- Crear la solución con módulos de interfaz, dominio, integración con Excel, servidor MCP, proveedores de IA y pruebas.
- Incorporar configuración, almacenamiento local y manejo de credenciales.
- Generar la plantilla, su perfil de reglas y casos de referencia con resultados esperados.
- Documentar instalación, configuración y ejecución del piloto.

**Entrega:** aplicación inicial que permite seleccionar la plantilla y consultar su estructura.

### Etapa 2: sesiones, copias y recuperación

- Trabajar con la versión guardada del archivo seleccionado. Crear una copia base inmutable y un borrador antes de abrirlo en la instancia dedicada de Excel.
- Mantener el original fuera de los destinos de escritura del ejecutor.
- Abrir el borrador sin actualizar vínculos ni ejecutar macros; rechazar archivos con capacidades excluidas.
- Crear un punto de recuperación antes de cada lote. Si una operación falla o deja el libro inválido, restaurar el lote completo.
- Registrar identificadores de operación para evitar duplicados tras reintentos o interrupciones.
- Detectar cambios manuales en el borrador. Antes de ejecutar o aprobar, verificar su revisión y descartar planes desactualizados.
- Recuperar sesiones interrumpidas desde el último punto válido, sin repetir escrituras automáticamente.

Este aislamiento se basa en copias y herramientas restringidas; no se presentará como una máquina virtual de seguridad.

**Entrega:** edición determinista del borrador con recuperación, todavía sin IA.

### Etapa 3: inspección, ejecución y validaciones

- Inventariar hojas, tablas, rangos, fórmulas, validaciones y relaciones del perfil.
- Resolver registros por identificadores, incluso después de ordenar o insertar filas.
- Implementar operaciones de datos y estructura con validación previa y posterior.
- Recalcular con Excel y comprobar errores, relaciones rotas, existencias y resultados del negocio.
- Calcular de forma independiente en C# los saldos y costos esenciales para detectar fórmulas incorrectas que no produzcan errores de Excel.
- Comparar valores, fórmulas, estructura y propiedades afectadas. Las referencias que no puedan analizarse con certeza se marcarán como no soportadas.

**Entrega:** herramientas MCP verificadas y comparación completa de cada lote.

### Etapa 4: agente conversacional y voz

- Conectar el modelo al inventario y a lecturas acotadas del libro. No enviar el archivo completo por defecto.
- Tratar el contenido de celdas como datos, sin permitir que cambie permisos o instrucciones del sistema.
- Convertir solicitudes en planes estructurados. Pedir aclaración cuando haya varios registros posibles, fechas ambiguas o campos obligatorios faltantes.
- Ejecutar automáticamente sobre el borrador las instrucciones suficientemente claras y permitidas.
- Añadir dictado, transcripción visible y entrada de texto equivalente. Ejecutar la instrucción al pulsar «Enviar».
- Mostrar estados comprensibles: escuchando, interpretando, modificando, comprobando y listo para revisar.
- Incorporar límites configurables de consumo, llamadas y duración. No conservar audio local después de transcribirlo; los registros técnicos omitirán datos de celdas y credenciales.

**Entrega:** flujo completo por voz o texto, con cambios visibles y explicación del resultado.

### Etapa 5: revisión, aprobación y distribución

- Mostrar cambios agrupados por instrucción: registros afectados, valores anteriores y nuevos, fórmulas modificadas, cambios estructurales y efectos calculados.
- Permitir corregir mediante otra instrucción, deshacer el último lote o descartar la sesión.
- Aprobar la revisión completa. No incluir aprobación parcial de celdas en esta versión.
- Vincular la aprobación a una copia inmutable exacta. Un cambio posterior invalidará la aprobación.
- Exportar a un archivo nuevo mediante selector de destino, sin sobrescribir el original ni otro archivo existente.
- Entregar aplicación Windows x64 autocontenida, plantilla, perfil, guía de uso e informe de pruebas. Excel seguirá siendo un requisito instalado.

**Entrega:** piloto utilizable de principio a fin.

## 5. Pruebas y criterios de aceptación

### Flujos funcionales

- Registrar una visita y sus materiales; verificar descuento y costo.
- Corregir una cantidad y observar actualización de inventario y resumen.
- Registrar una entrada y eliminar un consumo.
- Agregar una columna calculada, renombrar una hoja y conservar referencias.
- Deshacer un lote y exportar únicamente la revisión aprobada.

### Casos de protección y recuperación

- Existencias insuficientes, claves duplicadas, registros inexistentes y relaciones inválidas.
- Fórmulas con referencias rotas, ciclos o resultados contrarios a las reglas.
- Instrucciones ambiguas y texto malicioso dentro de celdas.
- Pérdida de conexión, Excel ocupado, cierre inesperado y reintento de una operación.
- Edición manual durante una sesión y cambios posteriores a la revisión.
- Intentos de escribir o exportar sobre el archivo original.

### Condiciones para aceptar el piloto

- El archivo original conserva su huella en todas las pruebas.
- Ningún archivo se exporta como aprobado sin una acción explícita del usuario.
- Todos los lotes fallidos recuperan un estado válido o dejan la sesión bloqueada para recuperación, sin declarar éxito.
- Los cálculos esenciales coinciden con el verificador independiente: identificadores y conteos exactos, cantidades con tolerancia de `0.000001` e importes de `0.01 MXN`.
- Una batería fija de 30 instrucciones cubre datos, fórmulas, estructura y ambigüedades. Al menos el 90 % de las instrucciones claras se resuelve correctamente; todos los casos ambiguos preparados solicitan aclaración y todos los casos inválidos se bloquean.
- Las cuatro hojas se revisan visualmente y el archivo aprobado se vuelve a abrir en Excel para verificar fórmulas, tablas y resultados.
- Las pruebas automáticas de dominio utilizan proveedores simulados; las pruebas reales de Excel y servicios de IA se ejecutan por separado en Windows, registrando versión del modelo, consumo y tiempos.

## 6. Referencias técnicas

Referencias consultadas durante la planificación del 29 de octubre de 2026. Verificar vigencia al implementar:

- [.NET: política de soporte](https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core)
- [Office: ejecución y modelo STA](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio)
- [Excel: Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas)
- [Excel: Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
- [Excel: Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
- [MCP: arquitectura](https://modelcontextprotocol.io/docs/learn/architecture)
- [Anthropic: modelos](https://platform.claude.com/docs/en/models/overview)
- [Anthropic: respuestas estructuradas](https://platform.claude.com/docs/en/build-with-claude/structured-outputs)
- [Azure Speech: idiomas admitidos](https://learn.microsoft.com/en-us/azure/ai-services/speech-service/language-support)
