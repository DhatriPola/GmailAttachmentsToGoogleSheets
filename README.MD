Gmail Attachments to Google Sheets

Problem Statement: 
The objective of this project is to automate the process of extracting the data from the excel attachement of an email to Google Spreadsheet and therefore build dynamic dashboards in looker studio.

Methodology:

![alt text](https://github.com/DhatriPola/GmailAttachmentsToGoogleSheets/blob/integration-setup/Architecture.png)


The first step of this is to create a label in your Gmail to get the emails from specific email everyday in one place. The next thing was to make the script configurable and for better user experience, I created a dialogue box to fetch the inputs: from which label should the data go to which sheet and then iterated through each email thread and each email message in the thread and I had to check if the email has any attachments.

Meanwhile, I figured out there is no direct way to pull the data from the email attachment to the google sheets. So, I decided to save the excel file to the drive temporarily and from the drive pull that data to the google sheet. After multiple updations to the code I was finally able to extract the data from the excel attachment to the google sheets.

Thereafter, I linked the google spreadsheet as the data source for the data studio to create dashboards and setup the refreshment of data every 12 hours. This way the dashboard automatically updates when the data adds up into the google spreadsheet everyday. 


