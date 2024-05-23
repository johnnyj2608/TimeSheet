## Project Name & Pitch

Timesheet

A Python script to automate the updating of time sheets for a new month.

2 packages were the core of this project: Xlwings and CustomTkinter

## How To Use

1. Press the "Select Folder" button to browse your file explorer and select a folder to edit all Excel files
2. Select a month and year you would like to modify all existing Excel sheets to
3. Click confirm changes and watch as a status bar underneath the button appears telling you how many have been modified so far

Executable will be located in dist directory

## Reflection

The project was born when for my current job, I had to monotonously create timesheets every month for every member/employee, which was over 500. 

I chose to use Xlwings over Openpyxl for editing my Excel sheets automatically because I found that Openpyxl does not handle dynamic data if I am using sequences or filter functions in Excel.

I was originally not going to include a GUI, but I wanted this script to be accessible to those with no coding skills. I researched ways to include a GUI and found that CustomTkinter was the best for the job based on how ease of designing.

Normally each timesheet would take 5 minutes to make by having a blank template of the month, copy-pasting personal information on top, and deleting rows that do not fit the individual's schedule. That easily was about 40 hours a month with the amount of timesheets necessary. Now that time is conserved to attend to more productive matters.

## Project Screen Shots

![Capture](https://github.com/johnnyj2608/TimeSheet/assets/54607786/5d9b4fdf-a32b-4dd4-a09a-a92d9d9ac5da)

