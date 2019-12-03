# process ProcXed reports

To prepare file:

* run ProcXed report (forthcoming publications)
* export as excel
* delete first row
* delete theme

Run script to process report.

# Moving to access format

After processing report, open csv file and:

* change any title flags as needed manually (get rid of edition)
* check for encoding issues
* save as .xls (Excel 97-03) for loading to Access
* rename old ForthcomingPubs.mdb file as ForthcomingPubs_[archive_date]
* open Access
* make blank db named ForthcomingPubs then: New > Import Table > Choose File > 
Select has headers > add to new Table > check it looks ok > 
add key automatically > choose name = "tblPubs"
* add rescheduled pubs are added as needed
* upload using FileZilla
* check website has updated

Line from github
