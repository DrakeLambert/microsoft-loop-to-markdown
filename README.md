# LoopToMd

Extracts content from Microsoft Loop HTML exports to markdown.

## Prerequisites

- .NET 9.0 SDK

## Getting HTML from Loop

1. Open your Loop page in a browser
2. Open Developer Tools (F12)
3. Go to Elements tab
4. Right-click on `<html>` element
5. Select "Copy" → "Copy outerHTML"
6. Save to a .html file

## Usage

```bash
dotnet run -- -i input.html -o output.md
```

## Options

- `-i, --input <file>` - Input HTML file (required)
- `-o, --output <file>` - Output markdown file (required)
- `-h, --help` - Show help

## Features

- Headings (h1-h3)
- Multi-level lists with indentation
- Checkboxes (checked/unchecked)
- Links in markdown format
- HTML entity decoding
