# 🏪 Store-Data-Collector-Robot

An automated Selenium-powered bot that scrapes store details from a dynamic web portal, processes multi-page data, and saves it neatly into Excel using OpenPyXL.

---

## 🚀 Features

- 🔐 **Automated login & navigation** — launches Chrome in headless mode and navigates securely through the store portal.  
- 🧭 **Dynamic element handling** — interacts with buttons, text fields, and next-page logic via Selenium.  
- 📄 **Scrapes detailed store data** — collects area, type, code, location, status, opening times, and contact information.  
- 📊 **Excel integration (OpenPyXL)** — writes structured data into Excel with formatting and sequential numbering.  
- 🔁 **Pagination automation** — automatically navigates through **all available pages** until the last page.  
- ⚙️ **Resilient architecture** — includes re-fetching, element checks, and smart delay handling for stability.  

---

## 🧠 Tech Stack

- **Python 3.x**
- **Selenium WebDriver**
- **OpenPyXL**
- **Headless Chrome (Options)**

---

## 🧩 How It Works

1. Loads target Excel file and initializes the Selenium WebDriver.  
2. Visits the main portal link and waits for store cards to load.  
3. Extracts store data fields one by one:
   - Area  
   - Store type and code  
   - Location  
   - Status (Open / Closed)  
   - Opening time details  
   - Contact & manager info  
4. Saves all data into the Excel file in real time.  
5. Continues through all pages until pagination ends.  

---

## 📁 Output Example

Each row in Excel contains:
| Area | Type | Code | Location | Status | Opening | Contact | Manager | Page | Store No. |
|------|------|------|----------|---------|----------|----------|----------|--------|------------|

---

## 🧑‍💻 Author

**Ahmed Essam**  
Automation Engineer | Python Developer  

---

## 🛠️ Future Enhancements

- Add screenshot capture for failed pages  
- Integrate progress logging system  
- Convert to `.exe` desktop tool for non-technical users  
