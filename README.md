# Excel ↔ Json Converter

**Author:** qiangxinglin

**Version:** 0.0.5

**Type:** tool

**Repository** https://github.com/qiangxinglin/excel-dify-plugin

## Description

The built-in `Doc Extractor` would convert input `.xlsx` file to markdown table **string** for downstream nodes (e.g. LLM). But this does not cover all situations! This plugin provides 2 tools:
- `xlsx → json`: Read the Excel file and output the Json presentation of the data.
- `json → xlsx`: Convert the given json string (list of records) to xlsx blob.



## Usage
> [!IMPORTANT]
> Correctly configure the **`FILES_URL`** in your `docker-compose.yaml` or [`.env`](https://github.com/langgenius/dify/blob/main/docker/.env.example#L48) in advance.

![](_assets/workflow_usage.png)

## Tools

### xlsx → json

- The output is placed in the `text` output field rather than the `json` field in order to preserving the header order.
- All cells are parsed as **string**, no matter what it is.
- If the uploaded Excel file contains multiple sheets, the plugin will automatically convert it into a JSON object, where each key is the sheet name and the value is the data array of the corresponding sheet.

| Name | Age | Date |
|------|-----|------|
| John |  18 |2020/2/20|
| Doe  |  20 |2020/2/2|


![](_assets/e2j_output.png)

### json → xlsx

- The output filename can be configured, default `Converted_Data`
- If the input JSON is an object (whose values are arrays), the plugin will automatically create a multi-sheet Excel file, where each key of the object will become a sheet name.

![](_assets/workflow_run.png)
![](_assets/output_xlsx.png)

#### Format Settings

The plugin supports optional formatting for row heights and column widths using the `[format]` reserved key.

> **Note:** Excel prohibits sheet names containing these characters: `/ \ ? * : [ ]`
> Therefore, `[format]` is guaranteed to never conflict with actual sheet names.

##### `[format]` Structure

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 20,           // Default height for all rows
      "columnWidth": 15,         // Default width for all columns
      "rowHeights": {            // Specific row heights (1-based)
        "1": 30,                 // Row 1 height = 30
        "2": 25                  // Row 2 height = 25
      },
      "columnWidths": {          // Specific column widths
        "A": 25,                 // Column A width = 25
        "B": 15                  // Column B width = 15
      }
    },
    "sheets": {
      "SheetName": {             // Per-sheet overrides
        "rowHeight": 22,
        "columnWidth": 18,
        "rowHeights": {"1": 28},
        "columnWidths": {"A": 30, "B": 20}
      }
    }
  },
  "SheetName": [...]             // Actual data
}
```

##### Format Priority Rules

Settings are applied in the following order (later overrides earlier):

1. `[format].defaults.rowHeight` / `columnWidth` - Global defaults for all rows/columns
2. `[format].sheets.<name>.rowHeight` / `columnWidth` - Per-sheet defaults
3. `[format].defaults.rowHeights` / `columnWidths` - Global specific rows/columns
4. `[format].sheets.<name>.rowHeights` / `columnWidths` - Per-sheet specific rows/columns

##### Validation and Warnings

- **Unknown sheet references**: If `[format].sheets` references a sheet that doesn't exist in the data, a warning will be displayed and those configurations will be ignored. The Excel file will still be generated successfully.
- **Type errors**: If format values have incorrect types (e.g., non-dict, negative numbers), an error will be thrown and the Excel generation will fail.

##### Examples

**Single sheet with global formatting:**

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 20,
      "columnWidth": 15
    }
  },
  "Sheet1": [
    {"Name": "John", "Age": "18"},
    {"Name": "Doe", "Age": "20"}
  ]
}
```

**Multiple sheets with per-sheet formatting:**

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 18
    },
    "sheets": {
      "Employees": {
        "columnWidths": {"A": 25, "B": 15},
        "rowHeights": {"1": 30}
      },
      "Departments": {
        "columnWidths": {"A": 20}
      }
    }
  },
  "Employees": [{"Name": "John", "Department": "R&D"}],
  "Departments": [{"ID": "1", "Name": "HR"}]
}
```

**Column identifiers:**

You can use either Excel letters or 1-based numeric indexes:

```json
{
  "[format]": {
    "defaults": {
      "columnWidths": {
        "A": 25,    // Letter format
        "1": 25,    // Numeric format (same as "A")
        "B": 15,
        "2": 15     // Same as "B"
      }
    }
  },
  "Sheet1": [...]
}
```


## Used Open sourced projects

- [pandas](https://github.com/pandas-dev/pandas), BSD 3-Clause License

## Changelog
- **0.0.5**: Add `[format]` metadata support for controlling row heights and column widths during JSON → Excel conversion
- **0.0.4**: Add missing dependency (xlrd)
- **0.0.3**: Add multi-sheet support for Excel processing (closes #13)

## License
- Apache License 2.0


## Privacy

This plugin collects no data.

All the file transformations are completed locally. NO data is transmitted to third-party services.