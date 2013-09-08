This is a civicrm extension to export directly into an excel format.
-------------

The goal is to free you from the csv->excel hell. The later has been proven guilty of the following charges:
- stubornly ignoring that unicode exists since 1992. 
- defaulting to the braindead latin1 (sorry, Windows-1252)
- doesn't offer the option to choose a sane char encoding
- and considering that the C in CSV stands for "Semi column" in excel in germany and france and ...
- being inconsistent with the US/UK version where C stands for Comma
- screwing up badly anytime there is a newline in one of the field
- doesn't offer the option to chose the column separator.
- in general made my life miserable

I would be paranoid, I'd think MS does it as a way to abuse its monopoly position and makes it as hard as possible to export/import to other software. 
But surely, it can't be possible that such a grown up software doesn't want to play nicely with its little software friends.
Surely.

Technical
--------------

The trick is to export an html table. with the right content type, it goes to excel and it import it smoothly

# Usage Instructions #
This extension affects the **Export Contacts** option for searches and group members, ie when you have done a search or listed the contacts in a group, you can choose what you want to do with the contacts that you have selected.  
This feature works exactly the same as it would otherwise, ie it takes you through the same 3 step process for selecting which fields to export but the file that it exports is in a different format.  
The file is exported with a XLS extension, so that when you open it, it will (normally) be opened in Excel but in fact the file is actually in HTML format, so when you open the file in Excel you will get a warning message saying (something like) ..  
> The file you are trying to open, ‘ *your_file_name*.xls’, is in a different file format than that specified by the extension.  Verify that the file is not corrupted and is from a trusted source before opening the file.  Do you want to open the file now?

.. and has buttons for **Yes**, **No** and **Help**.  
You can click **Yes** and the file will open.

However, when you come to save the file you will get a prompt (something like) ..  
> *your_file_name*.xls may contain features that are not compatible with Web Page.  Do you want to keep the workbook in this format?  
- To keep this format, which leaves out any incompatible features, click Yes.
- To preserve the features, click No.  Then save a copy in the latest Excel format.
- To see what might be lost, click Help.  

.. and has buttons for **Yes**, **No** and **Help**.  

You should click **No** and then you will get a prompt to save the file in the native Excel format (eg xlsx).  
Once the file is in the native Excel format you will not get these confusing messages again.
