# Excel Sidekick

An intelligent Excel companion that uses AI agents to help you understand complex risk management sheets.

## Overview

Excel Sidekick is designed for professionals working with massive, complex Excel workbooks—particularly in risk management and financial services. It acts like a dedicated expert who walks you through your sheets, explains calculations, traces dependencies, and helps you grasp the big picture.

### Key Features

- **Interactive Workbook Selection**: Automatically discover and select from all open Excel workbooks
- **Multiple Excel Instance Support**: Handles the same file open in different Excel.exe processes
- **Selection-Based Exploration**: Select any region in your Excel sheet and ask questions about it
- **Intelligent Dependency Tracing**: Automatically trace formula dependencies across sheets
- **Spatial Awareness**: Understands the layout and structure of complex multi-grid sheets
- **Semantic Annotations**: Add business terminology to ranges for better explanations
- **Flexible Configuration**: All paths support both relative and absolute paths
- **Cross-Sheet Analysis**: Seamlessly traces calculations across multiple worksheets

## Project Status

🚧 **Phase 1 - In Development**

Currently building the CLI-based foundation with selection-driven explanations.

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Run Excel Sidekick in interactive mode
python main.py

# Or use individual commands
python main.py list                       # List all open workbooks
python main.py connect                    # Connect interactively
python main.py connect "C:\Risk\VaR.xlsx" # Connect to specific file
python main.py build                      # Build dependency graph
python main.py ask "What does this calculate?"
python main.py explain                    # Explain current selection
python main.py trace Sheet1!A1 both 3     # Trace dependencies
```

### Interactive REPL Mode

```bash
# Start REPL
python main.py

# Available commands
excel-sidekick> list                      # List open workbooks
excel-sidekick> connect [full_path]       # Connect (interactive or by path)
excel-sidekick> build [--force]           # Build dependency graph
excel-sidekick> ask <question>            # Ask about workbook
excel-sidekick> explain                   # Explain selection
excel-sidekick> trace <cell> [dir] [depth]
excel-sidekick> annotate <range> <label> [desc]
excel-sidekick> search <query>
excel-sidekick> cache [status|rebuild|clear]
excel-sidekick> help                      # Show all commands
excel-sidekick> exit                      # Exit
```

## Architecture

Excel Sidekick follows a clean layered architecture:

- **Presentation Layer**: CLI interface (Web UI in Phase 2)
- **Application Layer**: Thin service facade
- **Domain Layer**: Core business logic with 5 services + LangGraph agent
- **Infrastructure Layer**: Excel connector, LLM providers, storage

## Configuration

All settings are managed in a single `config/config.yaml` file. Paths support both relative (from project root) and absolute paths. See the config file for all available options.

## Development

### Setup

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install development dependencies
pip install -r requirements-dev.txt

# Run tests
pytest

# Format code
black src/
ruff check src/
```

### Project Structure

```
excel-sidekick/
├── config/               # Configuration files
├── src/
│   ├── application/      # Application service layer
│   ├── domain/           # Core business logic
│   ├── infrastructure/   # External integrations
│   ├── presentation/     # CLI and Web API
│   └── shared/           # Shared utilities
├── tests/                # Test suite
└── docs/                 # Documentation
```

## Phase 1 Roadmap

- [x] Project structure and foundation
- [x] Excel integration (xlwings)
- [x] Dependency graph building and caching
- [x] Annotation system
- [x] Exploration agent (simplified for Phase 1)
- [x] CLI commands and REPL
- [x] Manual LLM provider (file-based)
- [x] Interactive workbook discovery and selection
- [x] Multiple Excel instance support
- [ ] Testing
- [ ] Documentation

## Future Phases

- **Phase 2**: Web UI, real-time monitoring, internal LLM API integration
- **Phase 3**: Automated sheet processing, knowledge base synthesis

## License

TBD

## Contributing

This is currently a personal project. Contribution guidelines coming soon.
