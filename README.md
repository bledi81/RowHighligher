# RowHighligher Excel Add-in

## Overview

RowHighligher is a Microsoft Excel Add-in built using Visual Studio Tools for Office (VSTO) and targets the .NET Framework 4.7.2. The add-in provides infrastructure to extend Excel with custom task panes and smart tags, enabling enhanced user interaction and automation capabilities.

## Features

- Integrates directly with Microsoft Excel via VSTO.
- Supports custom task panes for additional UI elements.
- Provides a framework for implementing smart tags.

## Project Structure

- `ThisAddIn.Designer.cs`: Auto-generated code that manages the add-in's lifecycle, initialization, and integration with Excel.
- `ThisAddIn.Designer.xml`: XML definition of the add-in's host items and controls.

## Requirements

- Visual Studio 2022 or later
- .NET Framework 4.7.2
- Microsoft Office Excel

## Getting Started

1. Clone the repository.
2. Open the solution in Visual Studio.
3. Build the project to restore dependencies.
4. Press F5 to run and debug the add-in in Excel.

## Notes

- The core logic and UI customizations should be implemented in the main add-in files (not the designer files).
- The designer files are auto-generated and should not be edited manually.

## License

Specify your license here.
