<div align="center">

# 🚀 GridPilot

**AI-assisted Windows desktop pilot for Microsoft Excel with bilingual UI, deterministic local execution, cloud LLM planning, and secure user-approved exports.**

![.NET 10.0](https://img.shields.io/badge/.NET-10.0-512BD4?logo=dotnet)
![Platform](https://img.shields.io/badge/Platform-Windows%20x64-blue)
![UI](https://img.shields.io/badge/UI-WPF-0078D4)
![MCP Server](https://img.shields.io/badge/Protocol-MCP-green)

</div>

---

## 📖 Overview

GridPilot is a Windows desktop application that automates and controls Microsoft 365 Excel workbooks using natural language, voice dictation, and structured batch operations. Designed with data safety as a primary requirement, GridPilot isolates working drafts from baseline copies and enforces explicit user verification before exporting any changes to disk.

The system combines an offline, deterministic C# interpreter with optional cloud planning via Google Gemini or Anthropic Claude. It also features a built-in Model Context Protocol (MCP) server for external AI agent integration, full SQLite transaction journaling, multi-level undo support, and real-time detection of manual spreadsheet edits.

## ✨ Features

- **Safe Draft-Based Workflow** — Creates an isolated working draft and read-only baseline for opened workbooks; original files are never modified directly.
- **Bilingual User Interface** — Dynamic, on-the-fly UI language switching between **English** and **Español** in both the main window and settings.
- **Hybrid Instruction Engine** — Supports deterministic offline parsing via `LocalInterpreter`, direct JSON batch payloads, and cloud LLM planning (Google Gemini `gemini-3.8-flash` or Anthropic `claude-sonnet-5-5`).
- **Voice Dictation** — Integrated push-to-talk speech recording (WinMM 16 kHz PCM) transcribed via Gemini Multimodal Audio.
- **Model Context Protocol (MCP)** — Embedded UTF-8 stdio MCP server (`GridPilot.Mcp`) providing 7 safe read/draft tools with automated orphan Excel process cleanup.
- **Transactional Audit & Recovery** — Durable SQLite journaling logs all modifications, enabling change comparison diffs, manual edit synchronization, and instant multi-step undo.
- **DPAPI Credential Protection** — API keys and user preferences are encrypted locally using Windows Data Protection API.

## 🛠️ Tech Stack

| Layer | Technology |
|:------|:-----------|
| **Language & SDK** | C# 13 / .NET 10.0 (`net10.0-windows`) |
| **User Interface** | WPF (Windows Presentation Foundation) |
| **Excel Interop** | Microsoft Office COM Interop (STA Thread, OLE Message Filter, OpenXML inspection) |
| **AI Providers** | Google Gemini API (`gemini-3.8-flash`), Anthropic API (`claude-sonnet-5-5`), Local C# Interpreter |
| **Audio Processing** | WinMM API (16 kHz PCM recording) + Gemini Multimodal Audio |
| **Storage & Persistence** | SQLite Journaling & Session Checkpoints |
| **Protocol** | Model Context Protocol (MCP) over UTF-8 stdio |

## ⚡ Quick Start

### Prerequisites

- **Operating System**: Windows x64
- **Software**: Desktop Microsoft 365 Excel installed
- **Development Tools** (for building from source): .NET SDK **10.0.403** (specified in `global.json`)

### Installation

Clone the repository and build the project using the PowerShell build script:

```powershell
git clone https://github.com/urielferia/grid-pilot.git
cd grid-pilot
./scripts/build.ps1 -Publish
```

The self-contained application bundle will be generated in `artifacts/release/`. No standalone .NET runtime installation is required to run the published build.

### Usage

1. **Launch Application**: Run the compiled binary:
   ```powershell
   ./artifacts/release/GridPilot.App.exe
   ```
2. **Select Language**: Toggle between **English** and **Español** in the top header or Settings (`Ajustes`).
3. **Open Workbook**: Click **Demo template** (`Plantilla demo`) or **Open workbook** (`Abrir libro`) to launch a draft instance in Excel.
4. **Execute Operations**:
   - **Local Batch**: Submit typed JSON operations into the instruction area:
     ```json
     [
       {"Kind":"SetValue","Table":"Movimientos","Id":"MOV011","Field":"Cantidad","Value":"3"}
     ]
     ```
   - **Natural Language / Voice**: Enter typed instructions or click **Speak** (`Hablar`) to dictate commands.
5. **Review & Reconcile**: Inspect cell diffs and calculated totals. Use **Refresh review** (`Actualizar revisión`) to capture manual edits made directly in Excel, or **Undo last batch** (`Deshacer último lote`) to revert changes.
6. **Approve & Export**: Click **Approve and export…** (`Aprobar y exportar…`) to save the finalized workbook to a new location.

## 📁 Project Structure

```
grid-pilot/
├── Directory.Build.props
├── GridPilot.sln
├── global.json
├── assets/
│   ├── GridPilot-Demo.xlsx
│   ├── expected.json
│   └── profile.json
├── docs/
│   ├── EVALUATION.md
│   ├── MODELS.md
│   ├── PROGRESS.md
│   ├── PROTOCOL.md
│   └── REFERENCES.md
├── scripts/
│   ├── build.ps1
│   └── test-mcp.ps1
├── src/
│   ├── GridPilot.App/           # WPF Desktop Client & Localization
│   ├── GridPilot.Core/          # Domain Contracts, Diffing & Rule Verification
│   ├── GridPilot.Excel/         # Excel COM Automation, STA Threading & OpenXML
│   ├── GridPilot.Mcp/           # Stdio MCP Server Engine
│   ├── GridPilot.Providers/     # Local Interpreter, Gemini & Anthropic Planners, Voice Input
│   └── GridPilot.Storage/       # SQLite Journaling, Checkpointing & Recovery
└── tests/
    └── GridPilot.Tests/         # Unit & Integration Test Suite
```

## 🔧 Configuration

All application settings are managed within the **Settings** (`Ajustes`) window:

- **AI Provider**: Choose between **Gemini** (`gemini-3.8-flash`) or **Anthropic** (`claude-sonnet-5-5`).
- **API Credentials**: Keys are entered in masked fields and encrypted locally via Windows DPAPI.
- **Cloud Consent**: Explicit opt-in toggle required before making outward requests to cloud provider APIs.

## 🧪 Testing

GridPilot includes unit tests, Excel interop verification, and MCP server smoke tests.

```powershell
# Set local CLI environment variables
$env:DOTNET_CLI_HOME="$PWD/.tools/cli"
$env:NUGET_PACKAGES="$PWD/.tools/packages"

# Run core unit tests
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release

# Run Excel COM integration tests (requires Excel installed)
./.tools/dotnet/dotnet.exe run --project tests/GridPilot.Tests -c Release -- --excel

# Run MCP server smoke test suite
./scripts/test-mcp.ps1 -Excel -Published
```

## 🤝 Contributing

Contributions to GridPilot are welcome! Please follow these guidelines:

1. Open an issue to discuss proposed changes prior to submitting a pull request.
2. Ensure all existing test suites (`GridPilot.Tests` and `test-mcp.ps1`) pass.
3. Maintain existing formatting standards and bilingual localization support.

## 📄 License

See source files for licensing details.