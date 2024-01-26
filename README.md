
# Search by Name (web scraping)

The objective of this project is to efficiently search and retrieve information (phone numbers, emails, previous adresses) on a large number of people using web scraping on [Fastpeoplesearch.com](https://www.fastpeoplesearch.com/).

## Requirements

- [Jupyter Notebook](https://jupyter.org/)
- [Python 3](https://www.python.org/)
- [Chrome](https://www.google.com/chrome/?brand=WDIF&ds_kid=43700078347700321&gclid=CjwKCAiAzc2tBhA6EiwArv-i6S09bM7LhhgfwroA4lCJRVVV0mykZ5vFpMPxV6WF1fgD7_zVSOEWkBoCxYMQAvD_BwE&gclsrc=aw.ds) (last update)

## How to Install and Run the Project

### Step 1: Open or Copy Files

- Open `fast-people-search.ipynb` in Jupyter Notebook.

- Alternatively, copy each code section (**A-J**) from `fps.py` to cells in Jupyter Notebook, using `fast-people-search.ipynb` as a reference.

### Step 2: Install Packages

- Uncomment Section **A** in the notebook.

- Run Section **A** to install the required packages using the package manager [pip](https://pip.pypa.io/en/stable/).

```bash
!pip install selenium
!pip install undetected-chromedriver
!pip install webdriver-manager
!pip install pandas
!pip install chromedriver-autoinstaller
```
## How to Use the Project

### Step 1: Prepare "input.xlsx" File:
- Download the input.xlsx file to the same folder as the project file with code.

- Open the `input.xlsx` file.

- Enter names, street addresses, and zip codes in Columns A, B, and C, respectively.

- Drag the fill handle down in Column D to apply the formula to all rows.

- Save changes made to `input.xlsx` and close the file.

### Step 2: Install the Project in Jupyter Notebook:

- Ensure the project is installed in your Jupyter Notebook environment.

- Execute the code in cells **B** to **F** in your Jupyter Notebook.

- Update the start and end lines in Section **G** based on your requirements.

- Execute cells **G** and **H** to initiate web scraping.

    ***Error Handling:***

    - If the code crashes, replace the start line number with the last number printed in cell **H** and rerun cells **G** and **H**.

    ***Handling Captcha:***

    - If a captcha appears, wait for 30 seconds before the search resumes.

    - Solve the captcha to continue web scraping.

    ***Reopen Browser (if needed):***

    - If required, reopen the browser by running cells **F** to **H**.

### Step 3: Save the results in "output.xlsx"

- Once cell **H** completes running, proceed to execute cells **I** and **J**.

- Open `output.xlsx` to see the results.

## License

Distributed under the [MIT License](https://choosealicense.com/licenses/mit/).

## Contact

Victoria Ten - victoria.ten.work@gmail.com