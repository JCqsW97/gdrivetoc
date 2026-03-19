# gdrivetoc - Google Drive Folder Table of Contents Generator

![Project Banner](https://images.unsplash.com/photo-1506744038136-46273834b3fb?ixlib=rb-4.0.0&auto=format&fit=crop&w=1600&q=80)

Create a comprehensive table of contents for any Google Drive folder with ease! This Google Apps Script allows you to generate a detailed list of all files and subfolders within a specified folder, directly inside a Google Spreadsheet.

---

## 📌 Features

- **Recursive Folder Listing:** Includes all files and subfolders within the chosen directory.
- **Customizable Output:** Easily format and customize the generated table, with nine different fields.
- **Easy to Use:** Attach the script to a blank spreadsheet for quick setup.
- **No External Dependencies:** Fully built with Google Apps Script, no additional tools required.

---

## 🚀 How It Works

1. **Setup:** Attach the script to a new or existing Google Spreadsheet.
2. **Run:** Generate the table of contents directly within your spreadsheet using the embedded menu UI.
3. **View:** Instantly see all files and subfolders listed with details like name, path, size, and URL.

---

## 📝 Usage Instructions

### 1. Attach the Script to Your Spreadsheet

- Open a new or existing Google Spreadsheet.
- Go to **Extensions > Apps Script**.
- Paste the script code (main.gs) into the editor.
- Create another file in the editor and paste the html code (FieldDialog.html).
- Save the project.

### 2. Run the Script

- Reload your spreadsheet.
- From the menu, select **Table of Contents** and select **Create a table of contents**.
- Choose the folder you want to generate a TOC thanks to its ID.
- Wait for the script to process and populate the sheet.

### 3. Customize & Automate

- From the menu, select **Table of Contents** and then **Add a field (Optional)**.
- From there, select a field to add to the TOC among the nine available.
- You can also modify the script to add more file details or automate the process via triggers.

---

## 📂 Sample Folder Structure

```
My Drive/
├── Folder1/
│   ├── FileA.docx
│   ├── FileB.xlsx
│   └── SubFolder/
│       └── FileC.pdf
└── Folder2/
    └── FileD.pptx
```

---

## 💡 Notes

- The script needs permission to access your Google Drive.
- For large folders, processing may take some time.
- Make sure the script is attached to a blank spreadsheet, at or near the root of your Drive for better usability.

---

## 🌟 Contribute & Feedback

Have suggestions or want to contribute? Feel free to open issues or pull requests!

---

## 📄 License

This project is licensed under the MIT License. See the [`LICENSE`](LICENSE) file for details.

---

*Happy organizing your Google Drive!* 🚀
