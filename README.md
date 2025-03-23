# Automated Lecture Scheduler from PDF Timetables

This project simplifies the process of scheduling recurring Google Meet lectures based on professors' timetables provided in PDF format. Using OCR and Google Calendar API, it extracts lecture details and schedules meetings for the entire semester, eliminating manual entry and ensuring accuracy.

## Features
- **OCR Processing**: Extracts timetable data from PDF files using `pdfplumber`.
- **Data Cleaning & Parsing**: Cleans and structures the timetable data using `pandas`.
- **Google Calendar Integration**: Uses Google Calendar API to schedule recurring meetings.
- **Automatic Email Generation**: Generates email IDs for each class based on extracted timetable information.
- **Fast & Accurate**: Schedules all lectures within **20 seconds** with **100% accuracy**.
- **User-Friendly Interface**: Simple GUI built with `PySimpleGUI`.

## How It Works
1. **Upload Timetable PDF**: Use the interface to upload the PDF timetable.
2. **Provide Semester Details**: Enter semester start date, number of weeks, and semester type (Even/Odd).
3. **Automatic Scheduling**: The program extracts lecture timings, generates email IDs, and schedules recurring Google Meet links.
4. **Confirmation**: Displays a success message once all events are created.
