# outlook_vba
VBA code to enhance the use of Outlook and mainly fight against spam, automatically assign emails to categories, ...

## How to install ?
1. Download the content of this repository on your hard disk (f.i. C:\Outlook_vba).
2. In Outlook, press ALT-F11 so the Visual Basic Editor (VBE) will be opened.
3. You'll see at the top left the explorer pane.  Right click under the root entry (first entry in the treeview) and select "Import" from the contextual menu.
* Import, one by one, the three files from the C:\Outlook_vba\VBA_CODE folder (i.e. files CAVO.bas, clsHelper.cls and  clsJSONLib.cls)
4. In the VBE, click on the Tools menu then select References and add a reference to "Windows Script Host Object Model" 

From here, the macro has been successfully registered, save it by, from the VBA, click on the File - Save (CTRL-S).  You'll probably be forced to give a name to this project.  Choose one.

You can close the VBA editor and go back to Outlook.

Now, add a button on the ribbon to be able to start the macro when you wish to :

1. Click on the File menu then select Options
2. In the Outlook Options window, click on "Customize Ribbon"
3. In the list "Choose commands from", select "Macros".  You'll see the "InspecteMails" macro
4. Select the macro name and click on the "Add >>" button to place the button where you wish (f.i. create a new group under "Home (Mail)"); just like you wish.
5. Click on the OK button when it's done.

Back in the Outlook interface, you should see the new buttons.

## How to use ?
If you've add a button in the ribbon, it's easy : open a folder with mails (f.i. the Inbox folder) and click on the button.   It isn't more difficult than that.

The macro will scan every emails in that folder (subfolders included) and will apply the defined rules.

## How to configure ?
Open your file explorer and go to the C:\Outlook_vba folder (where you've unzip this repository).  You'll find json files.  At time of writing this guide, there are two files.

Go to your "MyDocuments" folder (if you don't know where this folder is located, click on the Start button and type "%USERPROFILE%\My Documents\", Windows will then open that special folder) and create a folder called "outlook_vba".  Copy the two json files there.

When this is done; with a text editor like Notepad open these files.

###spam.json
This file contains the rules for identifying a spam email.

```
[
	"robots@altavista.com",
	"@aliyun.com",
	"@conexus.social",
	"@esab.co.uk",
	"@purple-office.com",
	"@revenue.com",
	"@sina.com",
	"@www13.jdays.net",
	"@zoho.com"
]
```

You can specify a full email address (like robots@altavista.com) or a domain name (like @zoho.com).
The InspecteMails macro will extract the sender email and will compare this info to this list.  If the sender email or sender domain name is mentionned in the spam.json file, the mail will be considered as a spam and will be deleted from your mail folder.

###categories.json
This file contains the rules for identifying a which Outlook category should be assigned to which emails.

```
{
    "one.friend@hotmail.com" : "Friends",
	"@aesecure.com": "aeSecure",
	"@afuj.fr": "Joomla!®",
	"@avonture.be": "Avonture",
	"@joomla.fr": "Joomla!®",
	"@paypal.be": "Paypal"
}
```

This file contains entries "email":"category".
The email can be a full email (like one.friend@hotmail.com) or just the domain name (like @aesecure.com).   
The category is an Outlook category (if the category doesn't exists yet, then it will be created).

The idea is : assign an Outlook category to emails received from these emails (or domain).  Outlook's categories allow to more easily group mails.  

See official documentation about categories in Outlook : https://support.office.com/en-us/article/Create-and-assign-color-categories-a1fde97e-15e1-4179-a1a0-8a91ef89b8dc

## Feedbacks

Don't hesitate to fork this repository and to share your changes with me.   
