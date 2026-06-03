# PPTX Editor LLM — Natural Language PowerPoint Editing via COM

## When to Use

Use this skill when you need to **edit an existing PPTX file using natural language instructions** on Windows with Microsoft Office installed. The LLM parses intent from Chinese or English instructions, and Windows COM automation executes the changes.

**Use this** (not python-pptx) when you need: animations, transitions, PDF export, image export, or any COM-exclusive feature.

## Prerequisites

- **Windows** with Microsoft Office (PowerPoint) installed
- **Python packages**: `pip install pywin32 requests`
- **API Key**: Set `OPENAI_API_KEY` environment variable (any OpenAI-compatible API)
- The PPTX file must exist on disk

## Architecture

```
User Instruction (natural language, CN/EN)
        ↓
   LLM Intent Parser (OpenAI-compatible API)
   - Receives PPTX structure (slides, elements, text, positions)
   - Outputs JSON array of structured actions
        ↓
   Action Dispatcher (pptx_editor_llm.py)
   - Validates and routes each action
        ↓
   PowerPointCOM (pptx_editor_com.py)
   - Executes via win32com against live PowerPoint instance
        ↓
   Save / Export
```

## Files

| File | Purpose |
|------|---------|
| `pptx_editor_com.py` | PowerPointCOM class with all COM methods + regex parser |
| `pptx_editor_llm.py` | LLM-based intent parser, replaces regex with AI understanding |

## Available Operations

| Category | Actions |
|----------|---------|
| Text | modify_text, modify_font, set_alignment |
| Appearance | set_fill, set_border |
| Layout | move_shape, resize_shape, delete |
| Add Elements | add_textbox, add_picture |
| Slides | add_slide, delete_slide, move_slide |
| Tables | modify_cell, table_row_add/delete, table_col_add/delete |
| Animation | animation (appear/fly/fade/zoom/bounce), remove_animation |
| Transitions | transition (fade/push/wipe/split/dissolve...) |
| Export | export_pdf, export_image |

## Usage Examples

### Single Instruction
```bash
python pptx_editor_llm.py presentation.pptx "把第一页标题改成红色加粗"
python pptx_editor_llm.py presentation.pptx "Add fade animation to slide 2 title"
```

### Interactive Mode
```bash
python pptx_editor_llm.py presentation.pptx --interactive
```

### Inspect Structure
```bash
python pptx_editor_llm.py presentation.pptx --inspect
```

### Dry Run (parse only)
```bash
python pptx_editor_llm.py presentation.pptx "删除第3页" --dry-run
```

### Batch Mode (pipe multiple instructions)
```bash
echo -e "把标题改成红色\n给第2页添加淡入动画" | python pptx_editor_llm.py presentation.pptx
```

### Custom API Endpoint
```bash
python pptx_editor_llm.py deck.pptx "export PDF" --api-base http://localhost:8080/v1 --model llama3
```

### As Python Module
```python
from pptx_editor_llm import run_single
run_single("deck.pptx", "把所有标题字号放大1.5倍", output="deck_modified.pptx")
```

## Configuration

| Env Variable | Default | Description |
|---|---|---|
| `OPENAI_API_KEY` | (required) | API key |
| `OPENAI_API_BASE` | `https://api.openai.com/v1` | API endpoint |
| `OPENAI_BASE_URL` | (fallback for above) | Alternative env var name |
| `OPENAI_MODEL` | `gpt-4o` | Model name |

CLI flags `--api-base`, `--model`, `--api-key` override env vars.

## ⚠️ Pitfalls

### BGR Colors (NOT RGB!)
COM uses BGR color format. Red = `0x0000FF` (255), Blue = `0xFF0000` (16711680).
Formula: `BGR = R + G*256 + B*65536`
The LLM system prompt includes a color reference table.

### 1-Based Indexing
All COM indices are 1-based: slides, rows, columns, characters, animation sequences.

### COM Lifecycle
- Always call `pythoncom.CoInitialize()` before COM operations
- Always call `.Close()` and `.Quit()` in finally blocks
- One PowerPoint instance at a time to avoid orphan processes

### STA Threading
COM requires Single-Threaded Apartment. If using from a thread, each thread needs its own `CoInitialize/CoUninitialize`. Don't share COM objects across threads.

### File Locking
PowerPoint locks the file while open. Don't open the same file in PowerPoint GUI and script simultaneously.

## Lessons from HTML-to-PPTX v2 Converter

### Object Classification
Different PowerPoint elements need fundamentally different handling. The LLM classifies intent into specific action types (text vs shape vs table vs animation) before execution.

### Fallback Strategy
When the LLM can't parse an instruction, it can:
1. Ask clarifying questions (interactive mode)
2. Make best guess based on context (batch mode)
3. Return null and skip (graceful degradation)

### Text Layout
COM handles text layout natively (unlike python-pptx which needs manual positioning). Use COM's TextFrame for all text operations — it handles wrapping, sizing, and paragraph formatting correctly.

### Batch Model
For multi-step edits, refresh the PPTX structure between each instruction to reflect cumulative changes. This prevents stale state issues (e.g., deleting slide 3 shifts slide 4 to position 3).
