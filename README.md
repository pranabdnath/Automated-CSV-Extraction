# Automated-CSV-Extraction-Filtering-and-Daily-Scheduler-Using-Google-Apps-Script
This script automates the process of extracting a CSV file from an email, unzipping it, filtering rows based on required conditions, removing duplicates, and updating a Google Sheet with only the latest valid records. Includes flexible date parsing and daily refresh logic.

Problem Statement 

The goal was to automate a recurring daily workflow that involved:

Receiving emails that contained compressed files (ZIP).

Extracting and reading a CSV file inside the ZIP.

Filtering rows based on predefined conditions.

Removing duplicate entries based on a unique key.

Selecting only records that match the current date.

Writing the cleaned and filtered data into a Google Sheet.

Ensuring old ZIP files inside a temporary working folder were cleared automatically.

The challenge was to convert a repetitive, manual task into an automated script that runs consistently every day without human involvement.

3. Why This Script Was Needed 

Before this automation:

Data had to be downloaded manually from an email.

The ZIP file had to be unzipped manually.

Multiple rows belonging to the same user caused duplication issues.

Data for previous days would mix with today's data unless manually cleaned.

Inconsistent date formats caused parsing errors.

Reports were delayed because the cleaning process took time.

This script solves all of these problems by handling downloading, extraction, filtering, cleaning, and updating the sheet automatically.

4. What the Script Does 
This solution:

Searches for the latest email using a subject placeholder.

Extracts the attached ZIP file.

Finds the CSV inside the ZIP.

Reads and parses the CSV.

Filters rows based on a custom set of conditions.

Removes rows with duplicate unique IDs.

Parses dates using a flexible date-reader function.

Selects only today's records.

Formats and writes the cleaned dataset into a target sheet.

Deletes the temporary ZIP file and clears any old leftover files.
