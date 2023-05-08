XLSX to PDF Converter App
This app reads data from two xlsx files, processes the data, and saves the results as PDF files. In order for the app to work, the two xlsx files must be in the same directory as the script.

Setup
Make sure you have Python installed on your system. If not, you can download it from the official Python website.

Create a Python virtual environment in the script's folder by running the following command:

Copy code
python -m venv venv
Activate the virtual environment:

On Windows:

Copy code
venv\Scripts\activate
On macOS and Linux:

bash
Copy code
source venv/bin/activate
Install the required libraries using the requirements.txt file:

Copy code
pip install -r requirements.txt
Configuration
Edit the paths.xlsx file to include the full path of the calculator xlsx file and the final path where you want the PDF files to be written to.

The source of truth xlsx file can be edited to add more data, but it must retain the exact format.

Running the App
To run the app, simply double-click the run.bat file. This will execute the app using the Python virtual environment you set up earlier.
