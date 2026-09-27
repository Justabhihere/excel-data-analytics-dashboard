# BITS Digital CodeForge V1.0 – Modern Grading Console

## Files
- `index.html` – page structure and UI
- `styles.css` – all visual styling and responsive layout
- `script.js` – Excel parsing, validation, analytics, grading, export, and interactions

## Run locally
Keep all three files in the same folder. Open `index.html` in a modern browser.

The app uses SheetJS from jsDelivr for `.xlsx/.xls` parsing, so an internet connection is needed when loading the library unless you later vendor the library locally.

## Expected Excel columns
- Student’s BITS ID
- Course
- Total Marks

Marks are expected on a 0–100 scale.

## Core flow
1. Enter instructor name.
2. Upload the Excel marks file.
3. Select a course.
4. Review analytics and student grades.
5. Adjust grade ranges if required.
6. Validate ranges.
7. Finalize and export CSV.

## CodeForge enhancements
- Stronger Excel/schema/data validation
- Duplicate BITS ID detection
- Searchable student table
- Analytics dashboard with distribution chart
- Empty-state handling
- Grade-range gap/overlap validation
- Safer CSV escaping
- Responsive modern UI
- Bug Fix Log and Enhancement Summary built into the application
