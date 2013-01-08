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
