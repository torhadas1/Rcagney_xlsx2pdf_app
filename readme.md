<title>XLSX to PDF Converter App</title>
<h1>XLSX to PDF Converter App</h1>
    <p>This app reads data from two xlsx files, processes the data, and saves the results as PDF files. In order for the app to work, the two xlsx files must be in the same directory as the script.</p>

    <h2>Setup</h2>

        <p>Make sure you have Python installed on your system. If not, you can download it from the <a href="https://www.python.org/downloads/">official Python website</a>.</p>

        <p>Create a Python virtual environment in the script's folder by running the following command:</p>
<pre><code>python -m venv venv</code></pre>

        <p>Activate the virtual environment:</p>
            <p>On Windows:</p>
<pre><code>venv\Scripts\activate</code></pre>

        <p>Install the required libraries using the <code>requirements.txt</code> file:</p>
<pre><code>pip install -r requirements.txt</code></pre>

    <h2>Configuration</h2>

        <p>Edit the <code>paths.xlsx</code> file to include the full path of the calculator xlsx file and the final path where you want the PDF files to be written to.</p>

        <p>The source of truth xlsx file can be edited to add more data, but it must retain the exact format.</p>


    <h2>Running the App</h2>

    <p>To run the app, simply double-click the <code>run.bat</code> file. This will execute the app using the Python virtual environment you set up earlier.</p>
