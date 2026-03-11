# RACI Matrix Builder

A comprehensive web-based application for creating, managing, and visualizing RACI (Responsible, Accountable, Consulted, Informed) matrices with support for multiple pages and various export formats.

![RACI Matrix Builder](screenshot.png)

## Features

### Core Functionality
- **Visual RACI Matrix Creation**: Intuitive interface for building RACI matrices
- **Multi-Project Support**: Work with multiple projects simultaneously and switch between them
- **Multi-Page Support**: Organize your matrices into multiple pages within a single project
- **Flexible Column Types**: Define columns as either "Stakeholders" or "Information"
- **Interactive Cell Editing**: Click on RACI badges to toggle assignments instantly
- **Real-time Updates**: Changes are saved automatically to browser localStorage

### Column Management
- **Stakeholder Columns**: Assign one or more RACI values (R, A, C, I)
- **Information Columns**: Store additional data and metadata
- **Column Descriptions**: Add descriptive text to columns for clarity
- **Easy Editing**: Quick access to edit or delete columns

### Row Management
- **Activity/Task Tracking**: Define activities or tasks as rows
- **Row Descriptions**: Add detailed descriptions to each row
- **Flexible Organization**: Add, edit, or delete rows as needed
- **Search & Filter**: Instantly search and filter rows by activity/task name
- **Row Sections**: Group related rows under sections for better organization
- **Bulk Operations**: Select multiple rows for batch actions like moving or deleting

### Export Options
- **JSON Export/Import**: Save and load projects in JSON format
- **Excel Export**: Export to .xlsx with each page as a separate sheet
- **PDF Export**: Generate PDF documents with professional formatting
- **PowerPoint Export**: Create .pptx presentations with each page as a slide

### Data Persistence
- **Auto-Save**: All changes are automatically saved to browser localStorage
- **Import/Export**: Transfer projects between browsers or share with team members

## What is RACI?

RACI is a responsibility assignment matrix that clarifies roles in project tasks:

- **R - Responsible**: The person who does the work to complete the task
- **A - Accountable**: The person who is accountable for the task being completed (only one per task)
- **C - Consulted**: People who provide input and are consulted during the task
- **I - Informed**: People who are kept informed about the task progress

## Getting Started

### Installation

1. Download or clone the RACI Matrix Builder files
2. Ensure you have the following files in the same directory:
   - `index.html`
   - `app.js`
   - `styles.css`

3. Open `index.html` in a modern web browser

**No server or installation required!** The application runs entirely in your browser.

### Browser Requirements

- Modern web browser (Chrome, Firefox, Safari, Edge)
- JavaScript enabled
- Internet connection (for CDN-hosted libraries)

## Usage Guide

### Creating a New Project

1. Open `index.html` in your browser
2. Enter a project name in the header
3. Start adding columns and rows to your matrix

### Managing Pages

**Add a Page:**
1. Click the "+" button in the Pages sidebar
2. Enter a page name
3. Click "Create"

**Switch Pages:**
- Click on any page name in the sidebar

**Delete a Page:**
1. Navigate to the page you want to delete
2. Click "Delete Page" button
3. Confirm the deletion

### Managing Columns

**Add a Column:**
1. Click "Add Column" button
2. Enter column name
3. Select column type (Stakeholder or Information)
4. Optionally add a description
5. Click "Save"

**Edit a Column:**
1. Click the edit icon (✏️) in the column header
2. Modify the details
3. Click "Save"

**Delete a Column:**
1. Click the edit icon in the column header
2. Click "Delete" button
3. Confirm the deletion

### Managing Rows

**Add a Row:**
1. Click "Add Row" button
2. Enter activity/task name
3. Optionally add a description
4. Click "Save"

**Edit a Row:**
1. Click the edit icon (✏️) in the row
2. Modify the details
3. Click "Save"

**Delete a Row:**
1. Click the edit icon in the row
2. Click "Delete" button
3. Confirm the deletion

### Searching and Filtering Rows

**Search for Rows:**
1. Use the search box above the table to filter rows
2. Type any part of an activity/task name
3. The table will instantly filter to show only matching rows
4. The results counter shows how many rows match (e.g., "Showing 5 of 20 rows")

**Clear Search:**
1. Click the "×" button next to the search box
2. Or clear the search text manually
3. All rows will be displayed again

**Tips:**
- Search is case-insensitive
- Partial matches are supported (e.g., "design" will match "Design Review", "UI Design", etc.)
- Section headers are preserved even when filtering
- Search works across all sections

### Assigning RACI Values

**For Stakeholder Columns:**
1. Click on the cell where the row and column intersect
2. Check/uncheck the RACI values (R, A, C, I)
3. Click "Save"

**For Information Columns:**
1. Click on the cell
2. Enter the information in the prompt
3. Click "OK"

### Exporting Your Matrix

**Save Project (Encrypted):**
1. Click "Save Project" button
2. An encrypted .raci file will be downloaded
3. This file is encrypted using AES-GCM encryption to protect your data

**Export as JSON (Plain Text):**
1. Click "Export" dropdown
2. Select "Export to JSON"
3. An unencrypted .json file will be downloaded (useful for debugging or integration)

**Export as Excel:****
1. Click "Export" dropdown
2. Select "Export to Excel"
3. The .xlsx file will be downloaded with each page as a separate sheet

**Export as PDF:**
1. Click "Export" dropdown
2. Select "Export to PDF"
3. The PDF will be generated with each page on a separate page

**Export as PowerPoint:**
1. Click "Export" dropdown
2. Select "Export to PowerPoint"
3. The .pptx file will be downloaded with each page as a separate slide

**Export as Word:**
1. Click "Export" dropdown
2. Select "Export to Word"
3. The .docx file will be downloaded with formatted tables and sections

### Importing a Project

1. Click "Import" button
2. Select a previously saved .raci file (or legacy .json file for backward compatibility)
3. If a project with the same name already exists, a dialog will appear with three options:
   - **Overwrite Existing Project** - Replaces the existing project (⚠️ permanently deletes existing data)
   - **Import as New Project** - Creates a new project with a numbered name (e.g., "Project (1)")
   - **Cancel Import** - Aborts the import without making any changes
4. The imported project will be added to your project list (or will replace an existing one if you chose overwrite)
5. Note: .raci files are encrypted to protect your RACI matrix data

### Managing Multiple Projects

**Switch Between Projects:**
1. Click on the project switcher dropdown (shows current project name)
2. Select any project from the list to switch to it
3. The active project is highlighted

**Create New Project:**
1. Click "New" button or select "New Project" from the project switcher dropdown
2. Enter a name for the new project
3. A new project will be created and added to your project list

**Close a Project:**
1. Click on the project switcher dropdown
2. Click the × (close) button next to any inactive project
3. Confirm to close the project (make sure it's exported if needed)
4. Note: Cannot close the last remaining project

## Data Structure

The application uses the following JSON structure:

```json
{
  "projectName": "My RACI Project",
  "pages": [
    {
      "id": "unique_id",
      "name": "Page 1",
      "columns": [
        {
          "id": "col_id",
          "name": "Project Manager",
          "type": "stakeholder",
          "description": "Leads the project"
        },
        {
          "id": "col_id_2",
          "name": "Priority",
          "type": "information",
          "description": "Task priority level"
        }
      ],
      "rows": [
        {
          "id": "row_id",
          "name": "Project Planning",
          "description": "Initial planning phase",
          "cells": {
            "col_id": ["R", "A"],
            "col_id_2": ["High"]
          }
        }
      ]
    }
  ]
}
```

## Technical Details

### Technologies Used

- **HTML5**: Semantic markup
- **CSS3**: Custom styling with Bootstrap 5
- **JavaScript (ES6+)**: Application logic
- **Bootstrap 5**: UI framework
- **Font Awesome**: Icons
- **SheetJS (xlsx)**: Excel export functionality
- **jsPDF**: PDF generation
- **jsPDF-AutoTable**: PDF table formatting
- **PptxGenJS**: PowerPoint generation

### Browser Storage

The application uses localStorage to persist your data:
- Data is stored in your browser
- No data is sent to any server
- Clearing browser data will remove saved projects
- Always export important projects as JSON for backup

### File Structure

```
RACIMan/
├── index.html      # Main HTML file
├── app.js          # Application logic
├── styles.css      # Custom styles
└── README.md       # Documentation
```

## Tips and Best Practices

### RACI Matrix Guidelines

1. **One Accountable (A) per task**: There should be only one person accountable for each activity
2. **At least one Responsible (R)**: Every task needs someone to do the work
3. **Avoid too many Consulted (C)**: Too many consultants can slow down decisions
4. **Keep Informed (I) relevant**: Only inform those who truly need to know

### Application Tips

1. **Use descriptive names**: Clear column and row names improve readability
2. **Leverage descriptions**: Add context to columns and rows for future reference
3. **Organize with pages**: Use pages to separate different project phases or workstreams
4. **Regular exports**: Export your project regularly as backup
5. **Column types**: Use Stakeholder columns for people/roles, Information columns for metadata like priority, status, etc.

## Keyboard Shortcuts

Currently, the application uses mouse-based interactions. Future versions may include keyboard shortcuts.

## Troubleshooting

### Data Not Saving
- Ensure browser localStorage is enabled
- Check that you're not in private/incognito mode
- Export your data as JSON regularly as backup

### Export Not Working
- Ensure you have a stable internet connection (required for CDN libraries)
- Check browser console for errors
- Try a different export format

### Display Issues
- Ensure JavaScript is enabled in your browser
- Try refreshing the page
- Clear browser cache and reload

## Future Enhancements

Potential features for future versions:
- Collaboration features
- Templates for common RACI matrices
- Advanced filtering and sorting
- Dark mode
- Keyboard shortcuts
- Undo/Redo functionality
- Copy/paste rows and columns
- Print optimization
- Cloud storage integration

## Contributing

This is an open-source project. Contributions, issues, and feature requests are welcome!

## License

This project is open source and available for use and modification.

## Support

For questions or support, please refer to the documentation above or create an issue in the project repository.

## Version History

### v1.0.0 (Current)
- Multi-project support with project switcher
- Smart import handling (conflict detection and resolution)
- Multi-page support within projects
- Section/grouping system for organizing rows
- Column type configuration (Stakeholder vs Information)
- One-click RACI toggling via colored badges
- Bulk operations (select, assign section, delete rows)
- Export to Excel, PDF, PowerPoint, Word, JSON
- Save/Load encrypted .raci files (with backward compatibility for .json)
- Dynamic column widths in exports (PDF/PPT/Word)
- localStorage persistence with multi-project storage
- Responsive design
- Row indentation for sections in all exports

---

**Built with ❤️ for better project management**
