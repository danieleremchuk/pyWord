# pyWord
Word Doc to CSV Converter

# What I learned
1. Python
2. Importing and using Python libraries
3. GUI

# Purpose of pyWord

I have another project in which I require the need to convert a Word document into a form better suited for a database. In this database, I will have records for each paragraph of the review, and tie the review back to an object. CSV import is the simplest way for me to import data into a database, in this case PostgreSQL. The idea to convert the Word document to a CSV file seemed to be a logical solution, and allowed me to try out a new language - Python.

# How it works

First, you select a Word document to read with the application. Then, you select a name (in this case, the brewery that is being reviewed) for your CSV file. Finally, you convert the file from Doc/Docx to CSV. The file will be placed in the directory where the application exists.

# Plans to enhance

I plan to add or enhance the following features (in no particular order):

- ~~CSV file preview~~ >> First iteration complete!
- Place file in a dedicated directory (separate from application directory)
- Integrate with the Database for a more direct import process (possibly without using a CSV file).
- Improve graphics/background/UI
