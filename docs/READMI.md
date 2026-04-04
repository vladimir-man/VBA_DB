# VBA_DB_Project

## description

Project for automating work with Excel and Word:

- Reading data from Excel tables by specified columns.
- Writing data to the DB file (new or selected).
- Exporting data from the "DB" into a formalized Word form.

## Structure

- `src/` — VBA source modules.
- `templates/` — templates of "DB" files (for example, BD_Template.xlsm).
- `docs/` - documentation and notes.

## Usage

1. Run the macro in Excel.
2. Select: create a new “DB” or open an existing one.
3. After filling in the data, click the **Start** button in the “DB” to export to Word.

## Target

Simplify your data management by collecting data from different Excel files and quickly generating reports in Word.

# Templates

## Purpose

This folder stores DB file templates used for recording data and exporting to Word.

## Structure

- `BD_Template.xlsm` — the main DB template:
- A sheet with a fixed-structure table (e.g.: ID, Date, Value1, Value2).
- A **Start** button for running the export macro.
- A VBA module with the procedure for exporting data to Word.

## Status

The `BD_Template.xlsm` file has not yet been created. It will be added in the next development stage.

## Purpose

The template ensures consistency in the DB structure and simplifies working with macros:

- All new DBs are created by copying this file.
- A pre-built button and macro are included.
