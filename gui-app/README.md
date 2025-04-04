# GUI Application for Excel Task Retrieval

This project is a GUI application that allows users to select an Excel file, retrieve task numbers from a specific column, and open a constructed URL in a web browser.

## Project Structure

```
gui-app
├── src
│   ├── main.py          # Entry point of the application
│   ├── gui
│   │   └── tabs.py      # GUI logic for the tabs
│   ├── utils
│   │   └── excel_handler.py # Utility functions for handling Excel files
├── requirements.txt      # Dependencies required for the project
└── README.md             # Documentation for the project
```

## Requirements

To run this application, you need to install the required dependencies. You can do this by running:

```
pip install -r requirements.txt
```

## Running the Application

To start the application, execute the following command in your terminal:

```
python src/main.py
```

## Features

- Select an Excel file using a file dialog.
- Display the selected file path.
- Retrieve task numbers from a specified column in the Excel file.
- Construct a URL using the retrieved task numbers and open it in a web browser.

## License

This project is licensed under the MIT License.