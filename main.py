# Copyright 2023 - Tom Smeets <tom@tsmeets.nl>
# Outlook ICS exporter, for sharing free/busy information with your other calendars
import win32com.client
import os

# OlDefaultFolders
# https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
olFolderCalendar = 9

# Calendar detail
olFullDetails         = 2
olFreeBusyAndSubject  = 1
olFreeBusyOnly        = 0

def export_outlook_calendar_to_ics(ics_file_path):
    outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get the default calendar folder
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getdefaultfolder
    folder = outlook_app.GetDefaultFolder(olFolderCalendar)

    # Export to ICS
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.calendarsharing.saveasical
    exporter = folder.GetCalendarExporter()
    exporter.CalendarDetail = olFreeBusyAndSubject
    exporter.IncludeAttachments = False
    exporter.IncludePrivateDetails = False
    exporter.IncludeWholeCalendar = True
    exporter.SaveAsICal(ics_file_path)

if __name__ == "__main__":
    # Store the .ics file relative to this main file
    cwd = os.path.dirname(os.path.abspath(__file__))
    ics_file_path = os.path.join(cwd, 'calendar.ics')

    export_outlook_calendar_to_ics(ics_file_path)
    print("exported to:", ics_file_path)
