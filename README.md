## How It Works

The script is designed to automatically retrieve store data from a web application using Selenium. It performs the following steps:

1. **open the intial web page and navigates** to the target section in the application.
2. **paginates through approximately 110 pages**, ensuring that all available store data is collected.
3. **Scrapes relevant fields** for each store (e.g., name, code, status, etc.).
4. **Writes the scraped data into an Excel file** using the `openpyxl` library.
5. **Generates a structured summary** that can be used for validation, reporting, or further processing.

This automation significantly reduces manual effort and ensures complete data capture across all available pages in the system.

## Author
Ahmed Essam
