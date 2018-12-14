# CONSEN
Edit word and convert to pdf and send email to the data provided in excel

Steps to follow:

Provide login details to send email:
	Provide user name, password
	Provide the SMTP/POP server name from which domain you are sending (servername: send.one.com and port number: 587(depends on the domain provider))
	Make sure to check “Check label” not to repeat the details. If the details repeat close and run the application again
Load the files in the order 
	Create two empty directories Like WORD_CREATED, PDF_FILES.
	Word File: The original word file with editable merge fields in the desired positions
How to add merge fields:
InsertExplore Quick parts  Fields  MergeFields  Add a Field Name
The mergefield name should be same as Column Name in Excel.
Press “Submit”
	Excel File: The excel with the mergefield data
Press”Submit”
	Directory Path: Load the path of the Directory (WORD_CREATED, PDF_FILES) where to save the updated Documents
Press “Submit”
	Sheet Number: Provide the sheet number of the date
Press”Submit”
	Make sure not to repeat the details
	Exit
