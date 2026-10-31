# GridPilot

Desktop application for Windows x64 and Microsoft 365 Excel. Operates on a visible draft workbook through typed operations, natural-language instructions (with a built-in deterministic local interpreter plus cloud planning via Gemini or Anthropic), and click-to-toggle voice dictation powered by Gemini Multimodal Audio. Exporting requires an explicit user approval over the complete review.

The user interface includes a built-in **Language** selector (**English** / **Español**) in both the main window header and the Settings dialog.

## Running

The self-contained distribution is built into `artifacts/release`. Run `GridPilot.App.exe` while keeping the `server` and `assets` folders next to the executable. Desktop Excel must be installed. No separate .NET installation is required to run the self-contained build.

Building from source requires .NET SDK **10.0.403** (`global.json`):

```powershell
./scripts/build.ps1 -Publish
./artifacts/release/GridPilot.App.exe
```

## Workflow

1. Select your preferred interface language (**English** or **Español**) at the top-right of the window or in **Settings** (`Ajustes`).
2. Click **Demo template** (`Plantilla demo`) or **Open workbook** (`Abrir libro`). GridPilot takes the saved version of the file, creates a read-only baseline and a working draft, and opens only the draft in a dedicated Excel instance.
3. Without cloud services, you can use **Local batch** (`Lote local`) with typed JSON operations, or write clear natural-language instructions (in English or Spanish) handled by the deterministic `LocalInterpreter`:

```json
[
  {"Kind":"SetValue","Table":"Movimientos","Id":"MOV011","Field":"Cantidad","Value":"3"}
]
```

4. For cloud planning and voice dictation, open **Settings** (`Ajustes`), enter your API key in the protected field, and enable cloud consent. New installations default to Gemini `gemini-3.8-flash`; legacy configurations preserve Anthropic `claude-sonnet-5-5`. Keys are protected via Windows DPAPI for the current Windows user outside the repository.
5. Type an instruction or click **Speak** (`Hablar`) to start recording (the button turns red and reads **Stop** / `Detener`) and click again when finished. Review the transcription and click **Send** (`Enviar`).
6. Review cell changes and summary totals in both windows. **Refresh review** (`Actualizar revisión`) detects manual edits in Excel. **Undo last batch** (`Deshacer último lote`) restores the previous checkpoint.
7. Click **Approve and export…** (`Aprobar y exportar…`), choose a new destination `.xlsx` file, and confirm. Existing files and the original workbook are never overwritten.

## Architecture

- `GridPilot.Core`: Contracts, structural/cell diffing, and independent domain rules verifier.
- `GridPilot.Excel`: Dedicated STA COM thread, bounded OLE message filter retries, package inspection, and OpenXML snapshots from native Excel copies.
- `GridPilot.Storage`: Durable SQLite journal, checkpoints, idempotency, crash recovery, explicit approval, and export.
- `GridPilot.Mcp`: Local JSON-RPC MCP server over UTF-8 stdio with orphan Excel process cleanup (`Owner.cs`). Exposes 7 read/draft tools and never exposes approval or export as model tools.
- `GridPilot.Providers`: Deterministic bilingual `LocalInterpreter`, `GeminiPlanner`, `AnthropicPlanner`, in-memory `SpeechInput` (WinMM 16 kHz PCM + Gemini Multimodal Audio), and DPAPI credential storage.
- `GridPilot.App`: WPF desktop client with dynamic English/Spanish UI localization.

## Tests

```powershell
$env:DOTNET_CLI_HOME="$PWD/.tools/cli"
$env:NUGET_PACKAGES="$PWD/.tools/packages"
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release -- --excel
./scripts/test-mcp.ps1 -Excel -Published
```
