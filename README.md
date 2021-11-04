Heavily modified from https://gallery.technet.microsoft.com/office/How-to-convert-Excel-xlsx-d9521619

Recurses through a folder that contains Excel files, converting each one to CSV and moving it as it is converted.  

Warning, this program is destructive on the original folder in a few ways, which may not be ideal--
it was what I needed for my use case but may not make sense for other folders.

In particular, the program as currently written ruthlessly removes printer settings from every excel file it encounters.
