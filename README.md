# PDF Index to Excel Converter 📚✨

This Python script is a powerful tool designed to extract structured data from PDF indexes (or similar tabulated text) and export it directly into an Excel spreadsheet. Leveraging the Google Gemini API, it intelligently parses terms and their corresponding page numbers, making it incredibly useful for researchers, data analysts, or anyone needing to digitize and organize index information.

-----

## Features ✨

  * **PDF Text Extraction:** Reads text content from specified PDF files using `PyPDF2`.
  * **AI-Powered Parsing:** Utilizes the **Google Gemini API** to accurately identify and separate "Term" and "Page Numbers" from raw index text.
  * **Structured Excel Output:** Organizes the parsed data into a clean, two-column Excel workbook (`.xlsx`) for easy analysis and management.
  * **Page-by-Page Processing:** Iterates through each page of the PDF, processing its content sequentially.
  * **User Feedback:** Provides real-time progress updates during PDF processing.

-----

## Prerequisites 🛠️

Before running the script, ensure you have the following installed:

  * **Python 3.x:** Download from [python.org](https://www.python.org/).
  * **Google Gemini API Key:** You'll need an API key to access the Gemini model. Get yours from the [Google AI Studio](https://aistudio.google.com/app/apikey).
  * **Required Python Libraries:**

-----

## Installation 💻

1.  **Clone the repository** (or download the `main.py` file directly if it's a standalone script):

    ```bash
    # If this script is part of a larger project, you'd clone the whole repo
    # git clone https://github.com/MemonDeveloper/PDF Index to Excel Converter.git
    # cd PDF Index to Excel Converter
    ```

2.  **Create a `requirements.txt` file:**
    In the directory where your script resides, create a file named `requirements.txt` and paste the following content:

    ```
    PyPDF2
    openpyxl
    google-generativeai
    ```

3.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

-----

## Gemini API Configuration 🔑

This script requires a Google Gemini API key to function.

1.  **Get your API Key:** Visit the [Google AI Studio](https://aistudio.google.com/app/apikey) to generate an API key.

2.  **Set the API Key in the script:**
    Locate the line:

    ```python
    genai.configure(api_key="YOUR_API")
    ```

    Replace `"YOUR_API"` with your actual Gemini API key.

    **Important:** For production or shared projects, consider using environment variables to store your API key instead of hardcoding it directly in the script for better security practices.

-----

## Usage ▶️

1.  **Place your PDF:** Ensure the PDF file you wish to process (e.g., `index_investor relations.pdf`) is in the same directory as your script, or update the `pdf_path` variable in the script to its correct location.

    ```python
    pdf_path = "index_investor relations.pdf" # Change this to your PDF file name
    ```

2.  **Run the script:**

    ```bash
    python main.py
    ```

    (Replace `main.py` with the actual name of your Python file, e.g., `pdf_parser.py` or `main.py`).

The script will begin processing page by page, and a confirmation message will appear once the Excel file is generated.

-----

## Output Format 📊

The script generates an Excel file (e.g., `index.xlsx`) in the same directory as the input PDF. The Excel file will contain two columns:

  * **Column A: Term**
  * **Column B: Page Numbers**

Each row will represent a unique term and its associated page references extracted from the PDF index.

-----

## Important Notes ⚠️

  * **PDF Format Expectation:** The script is optimized to parse text that resembles an index or glossary, where terms are followed by page numbers. Its performance may vary with different text structures.
  * **Gemini Prompt:** The parsing logic relies heavily on the prompt given to the Gemini model. If your PDF's index format is significantly different from the example provided in the prompt, you might need to adjust the `prompt` string within the `correct_with_gemini` function for optimal results.
  * **API Usage:** Be mindful of your Gemini API usage limits, especially when processing very large PDFs.

-----

## Support My Work ❤️

If you find this script useful and would like to support its development, any contribution is greatly appreciated\! Your support helps in maintaining and improving this tool.

Here are a few ways you can contribute financially:

  * **Patreon:** [Link to your Patreon page] (e.g., `https://www.patreon.com/YourUsername`)
  * **Buy Me a Coffee:** [Link to your Buy Me a Coffee page] (e.g., `https://www.buymeacoffee.com/YourUsername`)
  * **PayPal:** [Link to your PayPal.Me or direct PayPal donation link] (e.g., `https://paypal.me/YourUsername`)

Thank you for your support\!

-----

## Contributing 🤝

Feel free to fork this repository, open issues, and submit pull requests if you have suggestions for improvements or bug fixes. Ideas for future enhancements include:

  * **Command-Line Arguments:** Allow specifying input PDF path and output Excel path via command-line.
  * **Improved Error Handling:** More robust handling for PDF parsing errors, API call failures, or invalid input files.
  * **Batch Processing:** Add functionality to process multiple PDF files in a single run.
  * **Configuration File:** Implement a configuration file (e.g., YAML, JSON) for API keys and other settings.
  * **Support for Other Output Formats:** Option to export to CSV, JSON, or a database.
  * **More Flexible Prompting:** Allow users to customize the Gemini prompt for different PDF structures.
  * **GUI Interface:** Develop a simple graphical user interface for easier interaction.

-----
