# 📊 Smart Restaurant Processor

This Python script automates cleaning and enriching an Excel sheet of restaurant data. It removes duplicates, formats phone numbers, and generates Instagram search links based on restaurant names.

---

## 🚀 Features

* ✅ **Duplicate Removal**

  * Removes rows where both restaurant name and phone number are the same
  * Highlights partial duplicates (same name or same phone only) in **bold red**

* 📞 **Phone Number Formatting**

  * Ensures all numbers follow US format: `+1 (808)-XXX-XXXX`
  * Fixes common formatting issues and sets column to "Text" to avoid Excel auto-formatting

* 🔗 **Instagram Link Generation**

  * Uses Google Search syntax to find Instagram pages related to the restaurant name
  * Example: `https://www.google.com/search?q=Kono's+Northshore+hawaii+site:instagram.com`
  * Adds links as clickable hyperlinks in Excel

* 🧠 **Smart Menu Interface**

  * Press `Enter` to run all smart functions
  * Or select:

    * `[1]` Generate Instagram links only
    * `[2]` Remove duplicates only
    * `[3]` Improve phone number formatting only
    * `[0]` Export results to `output.xlsx`

* 🔄 **Overwrite-Safe Export**

  * Replaces existing `output.xlsx` if present

---

## 📁 Files

* `smart_restaurant_processor.py` – main script
* `input.xlsx` – input data (must be present)
* `output.xlsx` – generated result
* `requirements.txt` – dependencies

---

## 📦 Installation

```bash
pip install -r requirements.txt
```

### requirements.txt

```txt
pandas
openpyxl
tqdm
phonenumbers
```

---

## ▶️ Usage

Place `input.xlsx` in the same folder. Then run:

```bash
python smart_restaurant_processor.py
```

Follow the menu prompts to clean, enhance, and export your Excel data.

---

## 📌 Notes

* Instagram links are search-based for reliability
* Phone formatting assumes US numbers (Hawaii = 808)
* Script supports Windows and macOS/Linux terminal clearing

---

## 🧠 Suggestions or Issues?

Feel free to customize and extend this script to add:

* Verified Instagram scraping via SerpAPI or browser automation
* Export logs or summary reports
* Auto-open output on completion

---

© 2025 TableTurnerr | Hashaam Zahid