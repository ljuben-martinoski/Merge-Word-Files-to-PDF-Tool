import os
from docx import Document
from docxcompose.composer import Composer

# Step 1:telling the script whre the file are
# all file goes in the folder named 'documents'
folder_name = 'documents' # variable for the folder name

# this gets a list of all files in that folder
all_files = os.listdir(folder_name)

# this line makes suer we only look at word files (.docx)
#and sort them alphabetically(A to Z)
word_files = [] # creating a empty list to store the word files
# loop through all files in the folder
for f in all_files:
    if f.endswith('.docx'): # check if the file ends with .docx
        word_files.append(f) # add the file to the list 
# sort the list alphabetically
word_files.sort()

# Step 2: Start the Master dokument
# we open the very first file to start our combined document
first_file_path = os.path.join(folder_name, word_files[0]) # this gets the path to the first file
master_doc = Document(first_file_path) # variable master_doc in it Dokument cretes a document object form python-docx using the file path.
composer = Composer(master_doc) # variable composer creates a composer object that will help us add the other files to the master document

# Step 3:Loop through the rest of the files
#We skip the first file (since we already opned it) and add the others
# '[1:]' menas start from the second file in the list
for file_name in word_files[1:]:
    print(f"Adding file: {file_name}")

    # adding a page breack so each new file starts on a new page
    master_doc.add_page_break()

    # Open the next file 
    next_file_path = os.path.join(folder_name, file_name) # this gets the path to the next file
    next_doc = Document(next_file_path) # this opens the next file in the next_doc variable

    # stick it onto the end of the master documents
    composer.append(next_doc)


#  Step 4: Save the Big Word file
master_doc_name = "EVRITHING_MERGED.docx"
composer.save(master_doc_name)
prtint("Finishied merging! Saved as:", master_doc_name)


#Step 5: Turn it into a PDF using LibreOffice
prnt("Now converting to PDF....please wait.")


# This is like typing a command into your computer's terminal for you
# It tells LibreOffice (soffice) to turn the Word file into a PDF
os.system(f"soffice --headless --convertr-to pdf {master_doc_name}")
print("Done! You should see 'EVRITHING_MERGED.pdf' in the same folder.")







