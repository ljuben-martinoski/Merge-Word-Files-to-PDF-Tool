# Merge Word Files to PDF Tool

Ever had a bunch of Word documents that you needed to combine into one big file and then turn into a PDF? Yeah, me too. That's why I made this little script.

It takes all your Word files from a folder, merges them together (in alphabetical order, because why not?), and then converts the whole thing into a PDF. Simple as that.

## What It Does

- Takes all your `.docx` files from a folder and smushes them together into one document
- Sorts them alphabetically (A to Z) so you know what order they'll be in
- Adds a page break between each file so they don't run into each other
- Converts the final merged Word file into a PDF
- Does all of this automatically - just run it and go grab a coffee

## What You'll Need

Before you start, make sure you have:

1. **Python** - If you don't have it, grab it from [python.org](https://www.python.org/downloads/). Version 3.6 or newer should work fine.

2. **LibreOffice** - This is what converts the Word file to PDF. You can download it from [libreoffice.org](https://www.libreoffice.org/download/). 
   
   ⚠️ **Important**: After installing LibreOffice, you might need to add it to your PATH on Windows. If you get an error about `soffice` not being found, that's probably why. The usual location is `C:\Program Files\LibreOffice\program\soffice.exe`.

## Getting Started

First, get the code (if you haven't already):
```bash
git clone https://github.com/ljuben-martinoski/Merge-Word-Files-to-PDF-Tool.git
cd Merge-Word-Files-to-PDF-Tool
```

Then install the Python packages you need:
```bash
pip install python-docx docxcompose
```

That's it! You're ready to go.

## How to Use It

1. **Put your Word files in the right place**
   - All your `.docx` files go into the `MergeProject/documents/` folder
   - The script will find them and merge them in alphabetical order

2. **Run the script**
   ```bash
   cd MergeProject
   python mega_script.py
   ```

3. **Wait for the magic**
   - It'll print messages as it adds each file
   - When it's done, you'll have `EVRITHING_MERGED.docx` and `EVRITHING_MERGED.pdf` in the `MergeProject/` folder

## What's Going On Behind the Scenes?

Here's what the script does step by step:

1. Looks in the `documents/` folder and finds all the `.docx` files
2. Sorts them alphabetically (so "file1.docx" comes before "file2.docx")
3. Opens the first file and uses it as the starting point
4. Goes through all the other files one by one, adding them to the end with page breaks
5. Saves everything as `EVRITHING_MERGED.docx`
6. Uses LibreOffice to convert that Word file into a PDF

Pretty straightforward, right?

## When Things Go Wrong

### "soffice is not recognized"

This means your computer can't find LibreOffice. Try:
- Making sure LibreOffice is actually installed
- Adding it to your PATH, or
- On Windows, the full path is usually `C:\Program Files\LibreOffice\program\soffice.exe`

### "IndexError: list index out of range"

You probably don't have any `.docx` files in the `documents/` folder. Put at least one file in there and try again.

### "Permission denied"

The output file is probably open in Word or another program. Close it and run the script again.

### "ModuleNotFoundError: No module named 'docx'"

You forgot to install the packages! Run this:
```bash
pip install python-docx docxcompose
```

## Want to Change Something?

Feel free to tweak the script! Here are some easy things to modify:

- **Different folder?** Change `folder_name = 'documents'` to whatever folder you want
- **Different output name?** Change `master_doc_name = "EVRITHING_MERGED.docx"` to something else
- **Different sorting?** The sorting happens on line 20 - mess with that if you want files in a different order

## What This Uses

- `python-docx` - Lets Python read and write Word documents
- `docxcompose` - Does the actual merging of multiple Word files
- `LibreOffice` - Converts Word to PDF (you install this separately)

## About

Made by **Ljuben Martinoski** - [GitHub](https://github.com/ljuben-martinoski)

This is open source, so use it however you want. If you find bugs or want to add features, feel free to open an issue or send a pull request!

## Thanks To

- [python-docx](https://python-docx.readthedocs.io/) for making Word files easy to work with
- [LibreOffice](https://www.libreoffice.org/) for the PDF conversion magic
