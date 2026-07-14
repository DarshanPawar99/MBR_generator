# MBR Generator

This repository contains the first automation building blocks for client Monthly Business Review (MBR) presentations.

## Template layout

Original client presentations are preserved at the repository root. Working template copies are generated locally under `templates/` because binary `.pptx` files are not committed to PRs:

- `templates/telstra/Telstra_MBR_Template.pptx`
- `templates/clario/Clario_MBR_Template.pptx`

Client configuration files live under `configs/` and map the first nine MBR automation slides:

1. Reporting month and client name
2. Executive summary
3. KPI cards
4. Three-month sales trend
5. Current month vs previous month
6. Location performance
7. Restaurant performance
8. Feedback analysis
9. Business insights

## Local template setup

Run the setup utility in an environment where the original uploaded PowerPoint files exist at the repository root. It creates the ignored working template copies used by the client configs.

```bash
python scripts/setup_templates.py
```

## PowerPoint inspector

Run the inspector before building generation logic so important PowerPoint objects can be renamed and targeted safely by name rather than by shape index.

```bash
python src/pptx_inspector.py templates/telstra/Telstra_MBR_Template.pptx
python src/pptx_inspector.py --json templates/clario/Clario_MBR_Template.pptx
```

The report includes slide number, shape name, shape type, current text, chart type, table dimensions, image count, position and size in inches, and whether the element is likely editable.
