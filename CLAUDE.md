# PowerPoint Builder

This project generates polished .pptx presentations from natural language descriptions using `python-pptx`. The user describes what they want in Claude Code chat and gets an editable deck back.

## Workflow

### 1. Read the Style Guide First

Before generating any deck, read `references/style_guide.md` — it captures design preferences extracted from example slides (colors, fonts, layouts, diagram styles).

### 2. Analyze the Template

```bash
python -m markitdown templates/<name>.pptx       # read structure and text
python scripts/thumbnail.py templates/<name>.pptx  # visualize layouts
```

Extract the template's slide layouts, color scheme, and fonts. Use these when adding slides.

### 3. Generate the Presentation

Use `python-pptx` to:
- Open the template from `templates/`
- Add content slides using the template's existing slide layouts
- Build diagrams as **native PowerPoint shapes** (never embed diagrams as images)
- Save to `output/`

Import the diagram helpers:
```python
from diagram_helpers import (
    create_flowchart,
    create_process_flow,
    create_timeline,
    create_comparison,
    create_hierarchy,
    extract_template_colors,
)
```

### 4. Visual QA (Required Before Delivery)

```bash
python scripts/office/soffice.py --headless --convert-to pdf output/<name>.pptx
rm -f slide-*.jpg
pdftoppm -jpeg -r 150 output/<name>.pdf slide
ls -1 "$PWD"/slide-*.jpg
```

Inspect every slide image. Fix any overlapping elements, cut-off text, uneven spacing, or misalignment. Re-render and verify again after fixes.

### 5. Iterate

After delivering the first draft, ask the user what to change. Common requests:
- Resize/recolor diagram shapes
- Rearrange slide order
- Add/remove slides
- Change text content

## Diagram Rules

- All diagrams use native PowerPoint shapes — rectangles, diamonds, arrows, connectors, etc.
- Extract colors from the template's theme and apply them to shapes
- Use connector lines (`MSO_CONNECTOR_TYPE`) to link shapes
- Group related shapes so they move as a unit
- Standard spacing: 0.3" between elements, 0.5" from slide edges
- Keep diagrams simple and readable — prefer clarity over complexity

### Shape Reference

| Diagram Type | Shapes Used |
|---|---|
| Flowchart | Rectangles (process), diamonds (decision), rounded rectangles (start/end), arrows (flow) |
| Process Flow | Chevrons or rounded rectangles with arrows between steps |
| Timeline | Circles on a horizontal line with text above/below |
| Comparison | Side-by-side columns with headers and content blocks |
| Hierarchy / Org Chart | Rectangles with hierarchical connector lines |
| Concept Map | Rounded rectangles with labeled connector lines |

### Typography Defaults

| Element | Size | Style |
|---|---|---|
| Slide title | 36-44pt | Bold |
| Section header | 20-24pt | Bold |
| Body text | 14-16pt | Regular |
| Shape labels | 10-14pt | Regular or Bold |
| Captions | 10-12pt | Muted color |

## Project Structure

```
templates/     — .pptx template(s)
references/    — Example slides + style_guide.md
output/        — Generated presentations
diagram_helpers.py — Reusable diagram generation functions
```
