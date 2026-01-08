# ACC Roles Report Generator

This tool generates an interactive Sunburst chart and Role Builder report based on Excel data for Autodesk Construction Cloud (ACC) roles.

Access the live page [here](https://dodzikojo.github.io/hal-data-visualizations/output/ACC_Roles_Report.html#About).

## Overview

The application takes a spreadsheet of roles (`HAL - ACC Roles - Consenting - Generated Roles.xlsx`), processes the hierarchical relationships (Who/Party -> What/Function -> Where/Location), and outputs a static HTML report.

**Key Features:**
*   **Sunburst Visualization:** Interactive chart showing the distribution of roles.
*   **Role Builder:** Cascading dropdowns to explore valid role combinations.
*   **Search & Filter:** Find potential roles easily.
*   **Copy to Clipboard:** Quickly copy role names for use in other applications.

## Prerequisites

*   Python 3.7+
*   pip (Python package manager)

## Installation

1.  Clone the repository or download the source code.
2.  Install the required Python libraries:

```bash
pip install pandas plotly openpyxl
```

## Usage

1.  **Prepare Data:** Ensure the source Excel file is named `HAL - ACC Roles - Consenting - Generated Roles.xlsx` and is located in the root directory. The file must have a column named `Created Roles`.
2.  **Generate Report:** Run the Python script:

```bash
python generate_chart.py
```

3.  **View Results:** Open `output/ACC_Roles_Report.html` in your web browser.

## Technical Architecture

### `generate_chart.py`
This is the core processing script. It performs the following steps:
1.  **Data Ingestion:** Reads the source Excel file using `pandas`.
2.  **Parsing:** Splits the `Created Roles` column (e.g., "AP1 - Approver - Railhead") into its components:
    *   **Who (Level 1):** AP1
    *   **What (Level 2):** Approver
    *   **Where (Level 3):** Railhead
3.  **Hierarchy Extraction:** Builds a nested dictionary representing the valid relationships between levels (L1 -> L2 -> L3). This is crucial for the "Role Builder" cascading dropdowns.
4.  **Chart Generation:** Uses `plotly.express` to create an interactive Sunburst chart and saves it as `output/_chart_only.html`.
5.  **Report Integration:**
    *   Reads the existing `output/ACC_Roles_Report.html`.
    *   Updates the "Last Generated" timestamp.
    *   Injects the extracted hierarchy data directly into the HTML as a JavaScript object `const roleHierarchy = ...`.

### `ACC_Roles_Report.html`
The main user interface. It works as a standalone file but requires `_chart_only.html` in the same directory.
*   **Iframe Integration:** detailed chart interactions are handled safely using `window.postMessage` to communicate between the main report and the chart iframe, avoiding cross-origin security issues.
*   **Cascading Logic:** JavaScript functions (`getAvailableLevel2`, `populateRoleBuilderSelects`) use the injected `roleHierarchy` data to ensure users can only select valid combinations in the Role Builder.

## Files

*   `generate_chart.py`: Main execution script.
*   `output/ACC_Roles_Report.html`: The final report file (User Interface).
*   `output/_chart_only.html`: Selecting chart component (loaded via iframe).
*   `superseded/`: Archived scripts.
