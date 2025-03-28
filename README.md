# Excel Data Entry Application

This project is a simple web application that allows users to enter data into an Excel (.xlsx) file through a web interface. Users can specify the row and column where the data should be added.

## Project Structure

```
excel-data-entry
├── app.py                # Main application file
├── templates
│   └── index.html       # HTML form for user input
├── static
│   └── css
│       └── style.css    # CSS styles for the HTML page
├── requirements.txt      # Project dependencies
└── README.md             # Project documentation
```

## Requirements

To run this application, you need to have Python installed along with the following packages:

- Flask
- openpyxl

You can install the required packages using pip:

```
pip install -r requirements.txt
```

## Running the Application

1. Clone the repository or download the project files.
2. Navigate to the project directory.
3. Run the application using the following command:

```
python app.py
```

4. Open your web browser and go to `http://127.0.0.1:5000` to access the application.

## Usage

- Enter the names you want to add to the Excel file.
- Specify the row and column where the data should be inserted.
- Submit the form to write the data to the specified location in the Excel file.

## License

This project is open-source and available under the MIT License.