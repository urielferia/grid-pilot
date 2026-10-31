# Referencias verificadas durante la implementación (29/10/2026)

- .NET 10 LTS y versiones soportadas: https://dotnet.microsoft.com/en-us/platform/support/policy/dotnet-core
- Apertura de Excel, UpdateLinks y seguridad: https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
- Copias nativas: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas
- Detección de ciclos por hoja: https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.circularreference
- MCP base y mensajes JSON-RPC: https://modelcontextprotocol.io/specification/2025-06-18/basic
- Modelo Claude Sonnet 5.5 (`claude-sonnet-5-5`): https://platform.claude.com/docs/en/models/overview
- Herramientas estrictas y esquema estructurado: https://platform.claude.com/docs/en/build-with-claude/structured-outputs
- Azure Speech: https://learn.microsoft.com/en-us/azure/ai-services/speech-service/get-started-speech-to-text
- Microsoft.Data.Sqlite 10.0.14: https://www.nuget.org/packages/Microsoft.Data.Sqlite/10.0.14
- SQLitePCLRaw 3.0.5: https://www.nuget.org/packages/SQLitePCLRaw.bundle_e_sqlite3

La restauración detectó una vulnerabilidad en la dependencia transitiva SQLite 2.1.11 usada por Microsoft.Data.Sqlite 10.0.0. Se actualizó Microsoft.Data.Sqlite y se fijó explícitamente SQLitePCLRaw 3.0.5, sin desactivar la auditoría NuGet.

## Actualización 30/10/2026

- Gemini: https://ai.google.dev/gemini-api/docs/models/gemini-3.8-flash
- API HTTP y respuestas/firmas: https://ai.google.dev/api/generate-content
- Esquema JSON de herramientas: https://googleapis.github.io/js-genai/release_docs/interfaces/types.FunctionDeclaration.html
- Precios: https://ai.google.dev/gemini-api/docs/pricing
- JEV: https://docs.typesafe.ai/introduction y https://docs.typesafe.ai/introduction/quickstart

La recomendación de modelos y su costo ilustrativo están en MODELS.md. No se hicieron llamadas pagadas ni se afirma calidad de interpretación sin evaluación real.
