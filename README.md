
# Stock-Take Form Generator

This project was developed to produce printable stock-take checklists based on the database inventory for an electronic manufacturer. 

Includes mock data, example output and placeholder assets required to run.

## Screenshots

![App Screenshot](https://via.placeholder.com/468x300?text=App+Screenshot+Here)


## Detailed Process

Input:

- Reads a generic database inventory csv export
- Initialises a dataframe based on relevant columns
- Re-indexes columns to suit stock-take list
- Orders each row by the items location
- Adds an empty Quantity column to the end of the dataframe

Output:

- Declares the dataframe by calling the input functionality
- Initialises the docx file
- Sets the Document style and formatting
- Adds a header with a title, date and logo
- Initialises the document table, sets the format and style
- Declares the width of each column to suit data length
- Starts the main table building loops, adding appropriate data
- Reformats the description column font to save table space
- Finally, exports the the finished docx file. Option to export the dataframe used as a csv file
## Roadmap

- User input to declare which columns to use

- User input for formatting options

- Add a GUI

- Make executable with PyInstaller

- Improve csv data cleaning

- Add table to format docx header