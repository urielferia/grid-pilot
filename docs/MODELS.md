# Modelos y presupuesto de GridPilot

Recomendación del 30 de octubre de 2026: empezar con **Gemini 3.8 Flash** (`gemini-3.8-flash`) para el único agente de interpretación. Esta es una recomendación inicial basada en funciones y costo, no una afirmación de precisión medida: falta la evaluación real de 30 instrucciones.

## Qué modelo corresponde a cada componente

| Componente / herramienta | Modelo o implementación recomendada |
|---|---|
| Agente: interpretar español, resolver IDs, pedir aclaración, proponer operaciones y fórmulas | Gemini 3.8 Flash. Un solo modelo para toda la instrucción. |
| `read_records` | C# local; el agente Gemini solicita la búsqueda y razona sobre el resultado acotado. |
| `inspect_workbook` | C# + Excel COM; la interfaz obtiene estructura/perfil, sin otro LLM. |
| `read_range` | C# local; no requiere LLM. Actualmente no se expone al planificador de nube. |
| `propose_plan` | Salida estructurada del mismo Gemini; es un contrato del planificador, no una herramienta que escribe. |
| `execute_batch` | Controlador local + C#/COM. Recibe el plan validado; ningún LLM escribe directamente. |
| `recalculate_validate` | Motor de cálculo de Excel + verificador independiente C#. |
| `compare_revision` | Comparación determinista C#. |
| `restore_checkpoint` | Sesión/SQLite/archivos locales. |
| Revisión, aprobación y exportación | Interfaz y acción humana; ningún modelo autoriza exportaciones. |
| Dictado | Azure Speech, reconocimiento `es-MX` ya implementado; no usa la clave Gemini. Texto funciona sin Azure. |
| Explicación de cambios | Diferencias locales ya implementadas; no hace falta pagar otra llamada. |
| Clasificador opcional futuro | JEV de TypeSafe: elegir entre edición de datos, cambio estructural o aclaración. No integrado; no hace falta para este piloto. |
| Alternativa para comparar calidad | Anthropic `claude-sonnet-5-5`, seleccionado explícitamente y evaluado sobre los mismos 30 casos. |

No compraría crédito Anthropic todavía: primero mediría Gemini. Solo cambiaría si falla casos importantes o si Sonnet demuestra una mejora suficiente para justificar el costo. No hay fallback automático ni varios agentes cobrando por cada operación.

## Ejemplo de configuración

En Ajustes: proveedor `gemini`, modelo `gemini-3.8-flash`, clave **Gemini Developer API / AI Studio**, 4 llamadas máximas, 12 000 tokens observados por instrucción. Introducir secretos solo en los campos protegidos de la aplicación. Activar nube cuando se quiera empezar a consumir.

Las instalaciones nuevas sugieren Gemini; las configuraciones antiguas sin campo Provider conservan Anthropic y su modelo. Cambiar de proveedor muestra el modelo recomendado, pero solo Guardar confirma la selección. No hay sustitución automática ante errores.

La integración añadida usa `generativelanguage.googleapis.com/v1beta`, no Vertex AI. Un saldo de Google Cloud o un crédito promocional no debe darse por compatible sin comprobar el proyecto/cuenta de facturación. El usuario confirmó que el crédito está en AI Studio, compatible con este adaptador. Confirmó también JEV de TypeSafe AI.

## Ejemplo de costo con USD 5

La tarifa estándar publicada para Gemini 3.8 Flash al consultar es USD 0.75 por millón de tokens de entrada y USD 3.75 por millón de salida, incluyendo pensamiento, hasta el 31 de diciembre de 2026. Google anuncia USD 1.50 / 7.50 desde el 1 de enero de 2027.

Ejemplo hipotético por instrucción completa, sumando todas sus llamadas: 5 000 tokens de entrada y 1 000 de salida/pensamiento cuestan **USD 0.0075**. USD 5 alcanzarían aproximadamente 666 instrucciones de ese tamaño; la batería de 30 costaría alrededor de USD 0.225. No es una medición ni garantía: perfil, historial reenviado, lecturas, razonamiento y número de llamadas cambian el consumo. No incluye voz, impuestos ni otros servicios.

La app registra tokens y duración en usage.jsonl, sin instrucciones, celdas ni claves. El límite de tokens es por instrucción y se comprueba entre llamadas; no impide que una respuesta cruce el umbral ni constituye un tope monetario global. Configurar cuotas en la cuenta y empezar con unos pocos casos demo.

## Evaluación antes de aceptar el piloto

Seguir EVALUATION.md: al menos 18/20 instrucciones claras correctas, las 5 ambiguas deben pedir aclaración y las 5 inválidas deben bloquearse. Usar copias demo y registrar modelo exacto, tokens, tiempos y efectos. No confundir tests HTTP simulados con disponibilidad o calidad real del modelo.

## Fuentes oficiales

- [Catálogo Gemini y disponibilidad](https://ai.google.dev/gemini-api/docs/models)
- [Gemini 3.8 Flash](https://ai.google.dev/gemini-api/docs/models/gemini-3.8-flash)
- [Precios Gemini](https://ai.google.dev/gemini-api/docs/pricing)
- [API generateContent](https://ai.google.dev/api/generate-content)
- [FunctionDeclaration JSON Schema](https://googleapis.github.io/js-genai/release_docs/interfaces/types.FunctionDeclaration.html)
- [Claude Sonnet 5.5](https://platform.claude.com/docs/en/models/sonnet-5-5/overview)
- [Azure Speech: idiomas](https://learn.microsoft.com/en-us/azure/ai-services/speech-service/language-support)

## Dónde encaja JEV

La documentación oficial describe Choice, Score y Noul: decisiones entre opciones o sobre una rúbrica/afirmación. Para este proyecto podría clasificar la intención antes de llamar al planificador. Ejemplo: dado «agrega una columna de margen», seleccionar «cambio estructural». No generaría la fórmula ni la lista completa de operaciones, y su confianza no reemplazaría las validaciones ni la aprobación humana.

No añado esa llamada obligatoria al piloto: hay un solo agente y el enrutamiento extra aún no demuestra ahorro. Conservaría el crédito JEV para evaluar un clasificador opcional cuando haya suficientes solicitudes reales. El endpoint oficial es https://api.typesafe.ai/v1/systemone; la guía muestra jev-latest y una respuesta de ejemplo jev-1.13.0. Antes de integrar, fijar una versión disponible en la cuenta y medir falsos negativos; no usar un alias móvil para afirmar una evaluación reproducible.

Fuentes: [introducción oficial](https://docs.typesafe.ai/introduction) y [quick start/API](https://docs.typesafe.ai/introduction/quickstart).
