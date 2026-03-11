# PPT

A simple Python project for generating polished PowerPoint presentations (`.pptx`) programmatically.

## Features

- 16:9 widescreen slides
- Consistent colour theme (dark blue, mid blue, white)
- Pre-built slide types: Title, Agenda, Content (bullet list), Closing
- Easy to customise — edit `create_presentation.py` and re-run

## Requirements

```
python-pptx
```

Install with:

```bash
pip install python-pptx
```

## Usage

```bash
python3 create_presentation.py
```

This writes a `presentation.pptx` file in the current directory. Open it with
Microsoft PowerPoint, LibreOffice Impress, or Google Slides (File → Import).

## Customise

Open `create_presentation.py` and edit the slide-builder calls at the bottom of
`create_presentation()` to change titles, bullet text, colours, and the number
of slides.  Each helper function (`build_title_slide`, `build_agenda_slide`,
`build_content_slide`, `build_closing_slide`) can be reused or extended
independently.

## Output

| Slide | Description |
|-------|-------------|
| 1 | Title slide with full-bleed dark-blue background |
| 2 | Agenda (numbered list of sections) |
| 3 | Background & Context |
| 4 | Key Findings |
| 5 | Solution / Proposal |
| 6 | Next Steps |
| 7 | Closing / Thank You |
