## Codeforces Result Generator

This Python script is designed to **generate a table of results** from a specified Codeforces contest. It reads contestant handles from a Google Sheet and update the sheet with the results.

### Quick Start

#### 1. Setup the Environment

You need to install the required Python libraries.

* **Codeforces API Library:**
    ```bash
    pip install CodeforcesApiPy
    ```
    *(See [CodeforcesApiPy documentation](https://pypi.org/project/CodeforcesApiPy/) for more details.)*

* **Google Sheets Library:**
    ```bash
    pip install gspread-formatting
    ```

#### 2. Configuration

Create a file named `local_config.py` in your project's main folder. Example in `local_config.example.py`

You need to add the following information to the file:

* **Codeforces API Keys:**
    You can get your credentials here: [https://codeforces.com/settings/api](https://codeforces.com/settings/api).
* **Google Sheets Settings:**
    Details for accessing your Google Sheet (Service Account JSON file path, sheet ID, etc.).
* **Contest IDs:**
    The specific IDs of the Codeforces contests you want to analyze.

#### 3. Prepare the Google Sheet

The script needs a sheet named `handles` within your target Google Sheet.

This `handles` sheet must have **three columns** in this specific order:

1.  `surname`
2.  `name`
3.  `cf_login` (This must be the user's Codeforces handle)