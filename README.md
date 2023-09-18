https://developers.google.com/gmail/api/quickstart/python
 
>>Refer to the above link for enabling GMAIL API and Configuring OAuth consent screen.

>>Download and place the creadential file as "credential.json"

>>You can also use Firebase with this application to embed images in your emails.

>>In aaplication_configuration.json browser path for oauth screen and a preview email can be specified

>>Refer to the excel file for information regarding how the file should be formatted in order for application to work.



#defining variables in the email body file 

variables are defined like this #name# in the email body file. The variable should match the 
feild name in th excel file in order for it to work dynamically.

you can define as many as variables you like

refer test1.txt to know how variables are defined in the email body.

#embedding images in email body using firebase storage

You can also just send fancy posters or images in place of text in email. In order for that to work the image files
need to be uploaded on firebase stoarge then their link would be embedded in the email body.

refer to firebase_configuration.json for information regarding how it should be formatted. Also make sure that
suitable rule is in place in firebase stoarge in order for it to work.


#sending custom html as email body

You can also leverage HTML and CSS by creating a customised html file and sending it with email body.
