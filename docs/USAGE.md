# Excel Sidekick Usage Guide

## Overview

Excel Sidekick helps you understand complex Excel workbooks by using AI to explore and explain your sheets. This guide covers the basic workflow and commands.

## Prerequisites

1. **Windows with Excel installed** - Excel Sidekick uses xlwings to connect to Excel via COM automation
2. **Python 3.8+** with dependencies installed
3. **Open Excel workbook** - Have your Excel workbook open before connecting

## Getting Started

### 1. Installation

```bash
# Clone the repository
cd excel-sidekick

# Install dependencies
pip install -r requirements.txt
```

### 2. Configuration

All configuration is in [config/config.yaml](../config/config.yaml). Key settings:

- **Excel connection**: Window title pattern matching
- **Snapshot settings**: How to sample large ranges
- **Dependency tracing**: Maximum depth and direction defaults
- **LLM provider**: Manual file-based provider (Phase 1)
- **Logging**: Console and file output settings

## Dependency Modes

Excel Sidekick supports two dependency analysis modes. Choose based on your workbook size and needs.

### On-Demand Mode (Default)

**What it does:**
- Reads cells only when needed during dependency tracing
- No upfront graph building required
- Memory efficient

**When to use:**
- ✅ Large workbooks (50K+ formulas)
- ✅ Massive workbooks (100K+ formulas)
- ✅ Quick exploratory analysis
- ✅ Limited memory environments
- ✅ When you only need upstream (precedent) tracing

**Limitations:**
- Downstream (dependent) tracing not supported
- No full workbook impact analysis
- Each trace reads cells fresh (no caching)

**Configuration:**
```yaml
dependencies:
  mode: "on_demand"  # Default
```

### Full Graph Mode

**What it does:**
- Builds complete dependency graph upfront
- Uses row-based batching to prevent Excel crashes
- Caches graph for future sessions
- Enables full dependency analysis

**When to use:**
- ✅ Small to medium workbooks (<50K formulas)
- ✅ When you need downstream (dependent) tracing
- ✅ When you need impact analysis
- ✅ Repeated dependency queries on same workbook
- ✅ Complete dependency visualization

**How batching works:**
- Reads sheets in chunks (default: 1000 rows at a time)
- Prevents Excel COM crashes on large sheets
- Sequential processing (not parallel due to COM limitations)

**Configuration:**
```yaml
dependencies:
  mode: "full_graph"
  batch_size: 1000  # Rows per batch
```

**Building the graph:**
```bash
excel-sidekick> build
```

### Mode Comparison

| Feature | On-Demand | Full Graph |
|---------|-----------|------------|
| **Upfront build time** | None (instant) | Minutes for large workbooks |
| **Memory usage** | Low | High |
| **Upstream tracing** | ✅ Supported | ✅ Supported |
| **Downstream tracing** | ❌ Not supported | ✅ Supported |
| **Impact analysis** | ❌ Not supported | ✅ Supported |
| **100K+ formulas** | ✅ Safe | ✅ Safe (with batching) |
| **Trace performance** | Slower (reads cells) | Faster (uses cache) |

### Switching Modes

Edit `config/config.yaml`:

```yaml
dependencies:
  mode: "on_demand"   # or "full_graph"
  batch_size: 1000    # Only used in full_graph mode
```

Restart Excel Sidekick after changing modes.

### 3. Running Excel Sidekick

**Interactive REPL mode** (recommended):

```bash
python main.py
```

**Command-line mode** (single commands):

```bash
python main.py connect
python main.py ask "What does this sheet calculate?"
python main.py explain
```

## Basic Workflow

### Step 1: Discover Open Workbooks (NEW!)

**Discover all open workbooks across all Excel instances:**

```bash
excel-sidekick> discover
```

This shows:
- All open workbooks across all Excel.exe processes
- Process ID (PID) for each Excel instance
- Full file paths
- Number of sheets
- Unsaved changes indicator (*)
- Warning if same file is open in multiple instances

### Step 2: Connect to Workbook

**Option A: Interactive selection (recommended)**

```bash
excel-sidekick> connect
```

This will:
1. Discover all open workbooks
2. Display interactive table with all options
3. Let you select by number
4. Handle duplicates (same file in multiple Excel instances)
5. Prompt whether to build dependency graph

**Option B: Connect by full path**

```bash
excel-sidekick> connect C:\Risk\Models\VaR_Model.xlsx
```

If the file is open in multiple Excel instances, you'll be asked to select which one.

**What happens on connect:**
- Connects to Excel via xlwings using specific PID
- Loads workbook structure (sheets, cells, formulas)
- Prompts whether to build dependency graph (configurable)
- Displays workbook summary

### Step 3: Build Dependency Graph (Full Graph Mode Only)

**Note:** This step is only needed if you're using `full_graph` mode. In the default `on_demand` mode, the build command will inform you that no build is required.

If you're using full_graph mode and skipped graph building during connection, you can build it later:

```bash
excel-sidekick> build
```

Or force rebuild:

```bash
excel-sidekick> build --force
```

**What happens during build:**
- Reads sheets in batches (1000 rows at a time by default)
- Processes all formulas and builds dependency graph
- Saves graph to cache for future sessions
- Shows progress for large workbooks

**Build time estimates:**
- Small workbooks (<10K formulas): seconds
- Medium workbooks (10K-50K formulas): 1-2 minutes
- Large workbooks (50K-100K formulas): 2-5 minutes
- Very large workbooks (100K+ formulas): 5-10+ minutes

The dependency graph (full_graph mode) enables:
- Downstream dependency tracing
- Impact analysis
- Complete formula relationship visualization

### Step 4: Ask Questions

**Option A: Ask a general question**

```bash
excel-sidekick> ask How is Value at Risk calculated in this workbook?
```

**Option B: Explain a selection**

1. Select a range in Excel (e.g., A1:C10)
2. Run explain command:

```bash
excel-sidekick> explain
```

The agent will:
- Get your current Excel selection
- Expand context (nearby cells)
- Extract formulas
- Trace dependencies
- Query the LLM with gathered context

### Step 3: Interact with LLM (Manual Provider)

Since Phase 1 uses a manual file-based LLM provider, you'll see:

```
LLM query saved to: llm_input.txt

Please:
1. Open llm_input.txt
2. Copy the contents
3. Paste into your LLM (Copilot, internal LLMSuite, etc.)
4. Copy the response
5. Paste into llm_output.txt
6. Press Enter to continue
```

**Why manual?** This approach works in restricted corporate environments where API access may be limited.

### Step 4: Review Answer

The agent will:
- Read the LLM response from llm_output.txt
- Format it nicely with markdown rendering
- Display context used (if requested with `--show-context`)

## Commands Reference

### discover

Discover all open Excel workbooks across all instances.

```bash
discover
```

Shows:
- Process ID (PID) of each Excel instance
- Workbook name and full path
- Number of sheets
- Unsaved changes indicator
- Warnings for duplicates (same file in multiple instances)

Examples:
```bash
discover                         # Show all open workbooks
```

### connect

Connect to Excel workbook.

```bash
connect [full_path]
```

**Modes:**
- **Interactive** (no arguments): Shows discover, you select by number
- **By path**: Provide full path to specific workbook

Examples:
```bash
connect                                    # Interactive selection
connect C:\Risk\Models\VaR_Model.xlsx      # Connect to specific file
```

**Notes:**
- If same file is open in multiple Excel instances, you'll select which one
- Prompts whether to build dependency graph (configurable in config)
- Full paths are recommended over filenames to avoid ambiguity

### build

Build dependency graph for connected workbook (full_graph mode only).

```bash
build [--force]
```

**Mode behavior:**
- **On-demand mode**: Shows message that build is not required
- **Full graph mode**: Builds complete dependency graph with batching

Options:
- `--force, -f`: Force rebuild even if graph already exists

Examples:
```bash
build                        # Build graph if not exists (full_graph mode)
build --force                # Force rebuild (full_graph mode)
```

**When to use (full_graph mode only):**
- After connecting with graph building skipped
- After making significant formula changes
- To refresh stale cache
- When switching from on_demand to full_graph mode

**Performance:**
- Uses row-based batching (1000 rows per batch by default)
- Shows progress for large workbooks
- Build time depends on formula count (see Step 3 above)

### ask

Ask a question about the workbook.

```bash
ask <question>
```

Options:
- `--mode, -m`: Query mode (educational, technical, concise)
- `--show-context, -c`: Show context gathered

Examples:
```bash
ask What does this sheet calculate?
ask How is VaR computed? --mode technical
ask Explain the revenue calculation --show-context
```

### explain

Explain current Excel selection.

```bash
explain
```

Options:
- `--mode, -m`: Query mode (educational, technical, concise)
- `--show-context, -c`: Show context gathered

Examples:
```bash
explain                      # Explain current selection
explain --mode technical     # Technical explanation
```

### trace

Trace cell dependencies.

```bash
trace <cell_address> [direction] [depth]
```

Arguments:
- `cell_address`: Cell to trace from (e.g., Sheet1!A1)
- `direction`: upstream, downstream, both (default: both)
  - **upstream** (precedents): Cells that this cell depends on
  - **downstream** (dependents): Cells that depend on this cell
- `depth`: Maximum depth (default: 3)

**Mode limitations:**
- **On-demand mode**: Only upstream tracing supported
- **Full graph mode**: Both upstream and downstream tracing supported

Examples:
```bash
trace Sheet1!A1                    # Trace both directions (on-demand: upstream only)
trace Sheet1!B10 upstream 3        # Trace inputs only, depth 3
trace Summary!Z100 downstream 10   # Trace outputs, depth 10 (requires full_graph mode)
```

Output is a visual tree showing:
- Upstream dependencies (precedents/inputs to the cell)
- Downstream dependencies (dependents/cells that use this cell) - full_graph mode only
- Formulas and values at each node
- Depth level for each dependency

### annotate

Add or discover semantic annotations.

```bash
annotate [range] [label] [description]
annotate --discover [--sheet SHEET]
```

Examples:
```bash
annotate                                      # Discover all annotations
annotate --discover --sheet Summary               # List annotations for sheet
annotate Sheet1!A1:B10 "Revenue Inputs"       # Add annotation
annotate Sheet1!C1:C10 "Monthly Revenue" "Revenue by product line"
```

Annotations help the agent understand business context.

### search

Search annotations by text.

```bash
search <query> [--sheet SHEET]
```

Examples:
```bash
search revenue                    # Search all annotations
search VaR --sheet RiskMetrics    # Search in specific sheet
```

### cache

Manage dependency graph cache.

```bash
cache [status|rebuild|clear]
```

Examples:
```bash
cache                    # Show cache status
cache status             # Show cache status
cache rebuild            # Rebuild dependency graph
cache clear              # Clear cache
```

The cache stores the dependency graph to speed up subsequent connections. Rebuild if:
- You've made significant changes to formulas
- Cache seems stale
- Tracing shows incorrect dependencies

### status

Show connection and cache status.

```bash
status
```

Displays:
- Connected workbook
- Graph cache status
- Node and formula counts
- Annotations status

### help

Show help message with all commands.

```bash
help
```

## Query Modes

Excel Sidekick supports three query modes:

### educational (default)

Best for learning and understanding. Provides:
- Business context and purpose
- Step-by-step explanations
- Connections to financial concepts
- Examples

```bash
ask What does this calculate? --mode educational
```

### technical

For detailed technical analysis. Provides:
- Precise formula breakdowns
- Technical terminology
- Implementation details
- Edge cases

```bash
ask How is this computed? --mode technical
```

### concise

Brief, direct answers. Useful when you just need quick info.

```bash
ask What's the formula in A1? --mode concise
```

## Tips and Best Practices

### 1. Use Annotations Liberally

Add annotations for key regions:

```bash
annotate Sheet1!A1:A100 "Product IDs" "Unique identifier for each product"
annotate Sheet1!B1:B100 "Unit Prices" "Price per unit in GBP"
annotate Summary!Z100 "Total VaR" "Value at Risk at 95% confidence"
```

The agent uses these to provide better explanations.

### 2. Select Relevant Regions

When using `explain`, select:
- The output cell you want explained
- A few surrounding cells for context
- Don't select entire columns/rows (too much noise)

### 3. Ask Specific Questions

Good questions:
- "What does cell A10 calculate and why?"
- "How does this sheet compute Value at Risk?"
- "Where does the revenue figure in C50 come from?"

Vague questions:
- "Tell me about this sheet" (too broad)
- "Is this right?" (agent can't validate)

### 4. Use Trace for Complex Calculations

For deeply nested formulas:

```bash
trace Sheet1!A1 both 10
```

This shows the full dependency tree, helping you understand data flow.

### 5. Rebuild Cache After Major Changes

If you've:
- Added/removed many formulas
- Restructured the sheet
- Notice incorrect dependency traces

Then rebuild:

```bash
cache rebuild
```

### 6. Check Status Regularly

```bash
status
```

Shows you:
- Current connection
- Cache health
- Available annotations

## Troubleshooting

### "Not connected to a workbook"

**Solution**: Run `connect` first.

### "No selection found"

When running `explain`, you need to select a range in Excel first.

**Solution**:
1. Switch to Excel
2. Click and drag to select cells
3. Return to Excel Sidekick
4. Run `explain`

### "Connection failed: No Excel instance found"

Excel is not running.

**Solution**: Open Excel with your workbook, then try connecting again.

### Slow performance on large workbooks

**Solution**:
1. Use cache: `cache status` to verify it's enabled
2. Limit trace depth: `trace Sheet1!A1 both 3` instead of 10
3. Select smaller regions when using `explain`

### LLM response not working

**Solution**:
1. Check that llm_output.txt exists
2. Ensure you pasted the full LLM response
3. Check for formatting issues (plain text only)
4. Review llm_input.txt to see what context was provided

## Advanced Usage

### Custom Config

Create a custom config file:

```bash
cp config/config.yaml config/my-config.yaml
# Edit config/my-config.yaml
python main.py --config config/my-config.yaml
```

### Batch Mode (Command-Line)

Run multiple commands in a script:

```bash
#!/bin/bash
python main.py connect MyModel.xlsx
python main.py trace Sheet1!A1 both 5 > trace_output.txt
python main.py ask "Explain revenue calculation" > answer.txt
```

### Integration with Other Tools

Excel Sidekick can be integrated into workflows:

1. **Pre-meeting prep**: Annotate key regions, ask questions about calculations
2. **Model review**: Trace all outputs, verify dependencies
3. **Documentation**: Ask questions, save responses as model documentation

## Next Steps

- **Phase 2**: Web UI with real-time updates
- **Phase 3**: Automated sheet processing and knowledge synthesis

## Getting Help

For issues or questions:
- Check the [README](../README.md)
- Review [config/config.yaml](../config/config.yaml) for settings
- Look at code documentation in `src/`
