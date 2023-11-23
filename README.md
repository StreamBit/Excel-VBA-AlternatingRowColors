# Alternating Row Colors in Excel

This VBA script automatically alternates the background color of rows in an Excel worksheet based on changes in the values of a specified column. It is particularly useful for visually distinguishing adjacent rows that contain different data in a key column.

## Features

- Scans a specified column (default is Column B) for changes in values.
- Alternates the background color of rows every time a new value appears in the specified column.
- Easy to customize for different columns and color schemes.

## Getting Started

These instructions will guide you through the process of setting up and running the VBA script in your Excel workbook.

### Prerequisites

- Microsoft Excel
- Basic knowledge of Excel and its Developer tools

### Installation

1. **Enable Developer Tab** (if not already enabled):
   - Go to `File` > `Options`.
   - In the Excel Options dialog, select `Customize Ribbon`.
   - Check the `Developer` checkbox in the right pane.
   - Click `OK`.

2. **Open the VBA Editor**:
   - Click on the `Developer` tab.
   - Then click on `Visual Basic`, or press `Alt` + `F11`.

3. **Insert a New Module**:
   - In the VBA editor, right-click on any of the items in the `Project-VBAProject` pane.
   - Choose `Insert` > `Module`.

4. **Add the Script**:
   - Copy the provided VBA script into the new module.

### Usage

1. Open the Excel workbook where you want to use the script.
2. Run the macro by going to the `Developer` tab, clicking `Macros`, selecting `AlternateRowColors`, and then clicking `Run`.
3. The script will automatically alternate the row colors in the specified column (default is Column B).

### Customization

- To change the target column, modify the column reference in the script.
- To use different colors, change the RGB values in the script's color assignment lines.

## Author

StreamBit
