# Your Contact Merger

A simple tool to merge contacts from Google and MSSQL.

## Features

*   Merge contacts from Google Contacts CSV and MSSQL CSV files.
*   Handles different phone number formats and normalizes them.
*   Normalizes group names.
*   Provides a simple graphical user interface using `flet`.
*   Exports the merged contacts to a new CSV file.
*   Allows for a dry-run to see the results without saving the file.
*   Saves a log file with the merge details.

## How to use

1.  Install the dependencies from `requirements.txt`:
    ```
    pip install -r requirements.txt
    ```
2.  Run the application:
    ```
    python contacts_merger_Frontend.py
    ```
3.  Select your Google Contacts CSV file.
4.  Select one or more MSSQL contacts CSV files.
5.  Click on "Start Merge".

The merged contacts will be saved in a new CSV file in the `Contacts/output` directory.
