# Automatically Create Named Ranges Google Sheets
Created by [Perth SEO Agency](https://perth-seo-agency.com.au/)

**Automatically create named ranges in Google Sheets that can be updated dynamically.**

## How The Automatic Range Namer Works
This Google Apps Script adds a "Range Namer" menu to Google Sheets, enabling users to automatically create or update named ranges based on the headers of a selected data range. It sanitizes header names—replacing invalid characters and formatting them to comply with naming rules—and creates named ranges for each column excluding the header row. This script streamlines the process of maintaining valid named ranges, ensuring consistency and compliance within your spreadsheets.

### Features

-   **Automatic Named Range Creation**: Easily generate named ranges for each column in your selected data range using the first-row headers.
-   **Header Sanitization**: Ensures that all header names are formatted to comply with Google Sheets naming rules by replacing invalid characters and handling special cases.
-   **Update Existing Named Ranges**: Quickly update all named ranges after making changes to your headers by removing old named ranges and creating new ones based on the current selection.

### Installation

1.  Open your Google Sheets document.
2.  Navigate to `Extensions` > `Apps Script` to open the Apps Script editor.
3.  Delete any existing code in the editor and paste the provided script.
4.  Save the script by clicking the floppy disk icon or pressing `Ctrl + S`.
5.  Refresh your Google Sheet to see the new "Range Namer" menu appear.

### Usage

#### Auto Name Ranges

1.  Select the range of data you want to create named ranges for, ensuring the first row contains the headers and there is at least one additional row of data.
2.  Click on the `Range Namer` menu in the toolbar.
3.  Select `Auto Name Ranges`.
4.  The script will create named ranges for each column based on the sanitized header names.

#### Update Named Ranges

1.  After updating your headers or data range, select the new range.
2.  Click on the `Range Namer` menu.
3.  Select `Update Named Ranges`.
4.  This will remove all existing named ranges and recreate them based on your current selection.

### Header Sanitization Rules

-   **Dashes (`-`)** are replaced with underscores (`_`).
-   **Text inside brackets (`[]`)** is removed.
-   **Text inside quotation marks (`" "`)** is replaced with underscores. For example, `rel="next" 1` becomes `rel_next_1`.
-   **Special Characters**: All non-alphanumeric characters (except underscores) are removed.
-   **Leading/Trailing Underscores**: Any underscores at the beginning or end of the name are trimmed.
-   **Invalid Names**: If a header name is empty or invalid after sanitization, it is skipped.

### Important Notes

-   The script operates on the active sheet and uses the currently selected range.
-   Named ranges are created from the second row to the end of the selection, excluding the header row.
-   Ensure that your selection includes at least two rows: one for headers and one or more for data.
-   The script checks for valid names after sanitization and skips any headers that result in an empty name.

### Troubleshooting

-   **No Named Ranges Created**: Ensure that your headers are not empty and that they result in valid names after sanitization.
-   **Script Permissions**: The first time you run the script, you may need to authorize it to access your Google Sheets. Follow the prompts to allow the script to run.
-   **Menu Not Showing Up**: If the "Range Namer" menu doesn't appear, try refreshing the Google Sheet or ensuring the script is saved correctly.

### Contribution

Contributions are welcome! If you have ideas for improvements or encounter any issues, feel free to open an issue or submit a pull request on the GitHub repository.

### License

This project is licensed under the MIT License.





