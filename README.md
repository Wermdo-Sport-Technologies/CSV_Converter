# CSV Converter

A free CSV conversion utility by Wermdö Sport Technologies.

CSV Converter is a lightweight utility, not a full product suite.

It is primarily intended for extracting Taekwondo, Karate, or similar martial
arts grades from a grading registration form and exporting the names into
separate files for each grade.

## How the App Works

1. Choose a `.csv` file exported from a registration form.
2. The app reads the first row of the CSV as column headers.
3. Select the columns you want to include in the export, such as name, club, or
   grade.
4. Optionally enable duplicate removal.
5. Choose the column to split by, usually the grade column.
6. Choose an export format:
   - `xlsx` for Excel files
   - `md` for Markdown tables
   - `txt` for simple text files
7. Choose an output folder and run the conversion.

When a split column is selected, the app creates one output file per unique
value in that column. For example, if the selected split column contains grades
such as `8 kyu`, `7 kyu`, and `6 kyu`, the output folder will contain separate
files for those grades.

## Typical Use Case

The intended workflow is to take a grading registration CSV and quickly produce
cleaner grade-based lists. This can be useful when preparing grading groups,
printing checklists, or sharing names with instructors for each belt or grade
level.

## Limitations

- The app only reads CSV input files.
- The CSV must include a header row.
- It does not automatically detect which column is the grade column; you choose
  that manually.
- It does not validate martial arts grades or registration data.
- Duplicate removal compares full source rows, not only the selected export
  columns.
- Empty split values are exported as `Unknown`.
- Output files with the same name in the selected folder may be overwritten.
- The app is a Windows desktop app built with WPF and targets `.NET 8`.

## Notes

This tool is meant to solve a small, practical conversion task. It should work
well for simple grading registration exports, but more advanced reporting,
custom templates, or complex form cleanup may need extra manual work before or
after conversion.

## License

This software is free to use but may not be sold, rebranded, or claimed as another party's work.

See LICENSE for details.
