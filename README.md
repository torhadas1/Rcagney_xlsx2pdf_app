<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX to PDF Converter App</title>
</head>
<body>
    <h1>XLSX to PDF Converter App</h1>

    <p>This app reads data from two xlsx files, processes the data, and saves the results as PDF files. In order for the app to work, the two xlsx files must be in the same directory as the script.</p>

    <h2>Setup</h2>

    <ol>
        <li>Make sure you have Python installed on your system. If not, you can download it from the <a href="https://www.python.org/downloads/">official Python website</a>.</li>

        <li>Create a Python virtual environment in the script's folder by running the following command:</li>
<pre><code>python -m venv venv</code></pre>

        <li>Activate the virtual environment:</li>
        <ul>
            <li>On Windows:</li>
<pre><code>venv\Scripts\activate</code></pre>

            <li>On macOS and Linux:</li>
<pre><code>source venv/bin/activate</code></pre>
        </ul>

        <li>Install the required libraries using the <code>requirements.txt</code> file:</li>
<pre><code>pip install -r requirements.txt</code></pre>
    </ol>

    <h2>Configuration</h2>

    <ol>
        <li>Edit the <code>paths.xlsx</code> file to include the full path of the calculator xlsx file and the final path where you want the PDF files to be written to.</li>

        <li>The source of truth xlsx file can be edited to add more data, but it must retain the exact format.</li>
    </ol>

    <h2>Running the App</h2>

    <p>To run the app, simply double-click the <code>run.bat</code> file. This will execute the app using the Python virtual environment you set up earlier.</p>
</body>
</html>
