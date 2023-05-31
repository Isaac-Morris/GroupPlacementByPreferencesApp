# Group Placement By Preferences
A short program which takes two inputs- a list of places with minimum and maximum possible numbers of people, and a list of signups in small groups (1-3 people) and their preferred place options. This code is free to be used and added to under the GNU Affero General Public License v3.0.
Created by Isaac Morris 31/05/23

# How to use the Group Generator v2, with data from a Google Forms signup form:

1. Install Python from https://www.python.org/downloads/

2. After this has installed, open Command Prompt (on Windows) or Terminal (on Mac). Copy and press enter for these two lines of code:
"python -m ensurepip", and "python -m pip install pandas pathlib"

3. Download the data from the Google Forms signup form. You can do this by going to the form, then clicking
	Responses -> View in Sheets -> File -> Download -> Microsoft Excel (.xlsx)

4. Rename the file as signups_input.xlsx. Go into the file and rename the column headers (the first row) so they match the example file.

5. Create a file named places_input.xlsx. In this file, put the names of the places (EXACTLY AS THEY APPEARED IN THE SIGNUP FORM), as well as the minimum and maximum number of participants you want in each.
Note: If you're organising an event with students, I recommend adding a few students to the minimum numbers of each if possible- in our first Day of Good, 30-40% of signups didn't actually come, and half of these didn't even tell us they couldn't come lol

6. Double click the code! If it opens as a bunch of text, you may have to right-click on it and select "Run in Python".

The code will give you messages as it runs. These may tell you that there are problems in your data, and how to fix them. For example:
- There are places which people signed up for which aren't listed in the places_input files (you will have to replace these in one of the files so they match, probably by using Excel's Find and Replace function - CTRL-F)
- There are emails that appear in multiple signups (you will have to delete these manually)
- There are people who only signed up for places that don't exist
- There are people who only signed up for places which were already full of people who can't be moved somewhere else
- The code couldn't open the output file coz you currently have it open

If you want to see the code in action, try running it on the examples! There's a few errors in the data which you can figure out how to fix :D
