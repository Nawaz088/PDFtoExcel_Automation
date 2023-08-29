# PDFtoExcel_Automation
this project i have written an automated python script using selenium which takes PDF as an argument and returns Excel file

**Step-by-Step Guide: Running Excel Automation Script**

This guide will help you run an Excel automation script that processes data from a PDF file and creates a summary table in an Excel spreadsheet. No technical expertise is required; just follow these steps to set up and run the script on your machine.

**1. Prerequisites:**
- Ensure you have Python installed on your machine. If not, you can download and install it from the official [Python website](https://www.python.org/downloads/).

**2. Download ChromeDriver:**
- This script uses the Chrome web browser to interact with a website. Download the appropriate ChromeDriver for your version of Chrome from the [official website](https://sites.google.com/chromium.org/driver/).

**3. Install Required Libraries:**
- Open a terminal or command prompt on your computer.
- Run the following command to install the required Python libraries:

```
pip install selenium openpyxl pandas
```

**4. Save the Script:**
- Save the provided Python script to a folder on your computer. You can name the file `PDFtoExcel.py`.

**5. Configure ChromeDriver Path:**
- Open the `PDFtoExcel.py` script in a text editor.
- Replace the `chromedriver_path` variable with the path to the ChromeDriver executable you downloaded earlier. Make sure to use double backslashes (`\\`) in the file path.

**6. Run the Script:**
- Open a terminal or command prompt.
- Navigate to the folder where you saved the script using the `cd` command. For example:
  ```
  cd path\to\script\folder
  ```
- Run the script using the command:
  ```
  python PDFtoExcel.py
  ```

**7. Follow the Prompts:**
- The script will prompt you to select a PDF file using a file dialog.
- Once you select the PDF file, the script will automate the process of converting the PDF to Excel format and performing various data manipulations.

**8. Output:**
- The script will generate a modified Excel file named `Step4.xlsx` in the same folder where the script is located.
- This Excel file will contain the final processed data and a summary table.

**Note:**
- During the script execution, please do not interact with or open other applications, especially Excel. This can lead to permission errors while saving files.
- If you encounter any issues or errors, feel free to seek assistance or clarification.

Congratulations! You have successfully run the Excel automation script to process data from a PDF file and generate a summary table.
