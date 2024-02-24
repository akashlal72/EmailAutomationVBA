# EmailAutomationVBA
Sending email to multiple people with their respective document and customised subject and body line.
Automates Word Mail Merge for Individual Files

Problem Statement:

1.	Customization Requirement: Our business necessitates personalised documents for each customer, prompting the utilisation of mail merge for streamlined document creation.
2.	Drawback Identification: A notable challenge arises as the mail merge process produces a consolidated file, requiring additional steps for individual document extraction.
3.	Manual Separation: To address this, we currently resort to an external website for the manual separation of the merged document into individual files.
4.	Renaming Process: Subsequently, each file undergoes a renaming procedure for easy identification, adding an extra layer of manual effort.
5.	Time Consumption: In instances of extensive datasets, this manual intervention elongates the overall process duration, spanning approximately 3-4 hours.
6.	Error Prone: The manual nature of these tasks introduces a heightened risk of human errors, necessitating a more efficient and reliable solution for our document customization workflow.

Condition for running code:

Email 	CC 	Subject 	MsgText 	PicPath 	DocFolder 	PdfFolder 	FileName 	Type Of Attachment 	Contact 


1.	Keep the column's name and format as it is, as well as the order. Changing the order might lead to not working code.
2.	Run the code in by opening the word file VBA.
3.	Include these columns before needed before adding the dataset to the Word file.
4.	Make a unique docfolder and pdffolder for every document to avoid confusion for future purpose.
