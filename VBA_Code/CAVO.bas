Attribute VB_Name = "CAVO"
' ===========================================================================================
'
' Author : Christophe Avonture
' Date   : 17 August 2016
' Aim    : Provide quick functionnalities for
'
'   1. Kill mails coming from bad sender domain
'   2. Set email's category to predefined ones based on, here too, sender domain
'
' ===========================================================================================

Option Explicit
Option Base 1

Const cJSONFolder As String = "C:\Christophe\Repository\outlook_vba\"  ' By default, where to search for the JSON files

Const cKillInsteadOfDelete = False    ' Should spam emails be killed (TRUE) or just send to the deleted folder (FALSE)

Dim i As Integer, k As Integer
Dim JSONCategories As Object    ' Dictionnary object with the list of domains and the Outlook category that should be applied for these domains
Dim JSONSpam As Object          ' Dictionnary object with the list of domains that should be considered as spam
Dim JSONLib As New clsJSONLib   ' JSON parser
Dim cHelper As New clsHelper    ' Helper functions

' -----------------------------------------------------------------------------
'
' Initialization of variables and classes
'
' -----------------------------------------------------------------------------

Private Sub Initialize()

Dim sJSON As String, sFolder As String, sJSONFileName As String

   Set JSONLib = New clsJSONLib
   Set JSONCategories = Nothing
   Set JSONSpam = Nothing

   ' Preferred folder where to find JSON files.  If not found, try MyDocuments\OutlookVBA folder
   sFolder = cJSONFolder
   If (Right(sFolder, 1) <> "\") Then sFolder = sFolder & "\"
   
   If Not (cHelper.FolderExists(sFolder)) Then
      sFolder = cHelper.GetWindowsSpecialFolders("MyDocuments") & "OutlookVBA\"
      If Not (cHelper.FolderExists(sFolder)) Then
         cHelper.MakeFolder (sFolder)
         Debug.Print "Folder " & sFolder & " created"
      End If
   End If
  
   ' Get the list of domains considered as spam
   
   sJSONFileName = sFolder & "\spam.json"
   If cHelper.FileExists(sJSONFileName) Then
      sJSON = cHelper.GetFileContent(sJSONFileName)
      Set JSONSpam = JSONLib.parse(sJSON)
   Else
      Debug.Print "File " & sJSONFileName & " not found"
   End If
   
   ' Get the list of domains for which a specific category should be applied (a category in Outlook allow to group emails)
   
   sJSONFileName = sFolder & "\categories.json"
   If cHelper.FileExists(sJSONFileName) Then
      sJSON = cHelper.GetFileContent(sJSONFileName)
      Set JSONCategories = JSONLib.parse(sJSON)
   Else
      Debug.Print "File " & sJSONFileName & " not found"
   End If

End Sub

Private Sub Finalize()

   Set JSONSpam = Nothing
   Set JSONCategories = Nothing
   Set JSONLib = Nothing

End Sub

Private Sub processFolder(ByVal oParent As Outlook.MAPIFolder)

Dim oFolder As Outlook.MAPIFolder, oDeletedFolder As Outlook.MAPIFolder
Dim o As Object
Dim oMailItem As MailItem
Dim oProperty As Object
Dim sMailSenderAddress As String, sMailSenderDomain As String, sMailEntryID As String
Dim bContinue As Boolean, bHasBeenChanged As Boolean

   Debug.Print Replace(Space(20), " ", "=") & " Process folder " & oParent.Name & " " & Replace(Space(20), " ", "=")

   ' Loop every mails from that folder
   For Each o In oParent.Items

      ' Process only emails
      If TypeName(o) = "MailItem" Then

         bHasBeenChanged = False

         Set oMailItem = o

         sMailSenderAddress = oMailItem.SenderEmailAddress
         sMailSenderDomain = Mid(sMailAddress, InStrRev(sMailAddress, "@") + 1)

         bContinue = True

         Set oDeletedFolder = Application.Session.GetDefaultFolder(olFolderDeletedItems)

         k = JSONSpam.Count
         
         For i = 1 To k
            
            ' JSONSpam.Item is either a full email address like spammer@hotmail.ru or a domain like @hotmail.ru
            
            If ((sMailSenderAddress = JSONSpam.Item(i)) Or _
                (Left(JSONSpam.Item(i), 1) = "@") And (sMailSenderDomain = JSONSpam.Item(i))) Then
               Debug.Print "eMail from @" & sMailSenderDomain & " detected; kill it.  Mail subject was " & Chr(34) & oMailItem.Subject & Chr(34)
               oMailItem.UserProperties.Add "DeleteMeNow", olText
               oMailItem.Save
               oMailItem.Delete
               bContinue = False
               Exit For
            End If
         Next

         If bContinue Then

            k = JSONCategories.Count

            For i = 1 To k

               ' JSONCategories.keys()(i) = full email or domain of the mail sender (f.i. @hotmail.ru)
               ' JSONCategories.Items()(i) = categories to set for that domain

               If ((sMailSenderAddress = JSONCategories.keys()(i - 1)) Or _
                  (Left(JSONCategories.keys()(i - 1), 1) = "@") And (sMailSenderDomain = JSONCategories.keys()(i - 1))) Then
                  
                  If (oMailItem.Categories <> JSONCategories.Items()(i - 1)) Then
                     Debug.Print "Set category to " & Chr(34) & JSONCategories.Items()(i - 1) & Chr(34) & " for email coming from " & sMailSenderAddress
                  
                     oMailItem.Categories = JSONCategories.Items()(i - 1)
                     bHasBeenChanged = True
                     Exit For
                     
                  End If
               End If

            Next i

            If bHasBeenChanged Then
               On Error Resume Next
               oMailItem.Save
               If Err.Number <> 0 Then Err.Clear
               On Error GoTo 0
            End If

         End If ' If bContinue Then

      End If ' If TypeName(o) = "MailItem" Then

   Next o ' For Each o In oParent.Items

   ' Process subfolders if any
   If (oParent.Folders.Count > 0) Then
      For Each oFolder In oParent.Folders
         processFolder oFolder
      Next
   End If

   ' -----------------------------------------------------------------------------
   '
   ' Now, kill definitively mails that were deleted here above
   ' (otherwise stay in the Deleted folder
   '
   ' -----------------------------------------------------------------------------

   If cKillInsteadOfDelete Then

      Set oDeletedFolder = Application.Session.GetDefaultFolder(olFolderDeletedItems)

      For Each oMailItem In oDeletedFolder.Items
         Set oProperty = oMailItem.UserProperties.Find("DeleteMeNow")
         If TypeName(oProperty) <> "Nothing" Then oMailItem.Delete
      Next

      Set oProperty = Nothing
      Set oDeletedFolder = Nothing

   End If

   Set o = Nothing
   Set oFolder = Nothing
   Set oMailItem = Nothing
   
   Debug.Print Replace(Space(20), " ", "=") & " DONE " & Replace(Space(20), " ", "=")

End Sub

' ---------------------------------------------------------------------------------------
' -                                                                                     -
' - Entry point, a button in the Outlook interface can be added to fire this subroutine -
' -                                                                                     -
' ---------------------------------------------------------------------------------------

Sub InspectInbox()

   Call Initialize

   ' Process the current folder i.e. the active folder when the user has clicked on the InspectEmails button
   Call processFolder(Application.ActiveExplorer.CurrentFolder)

   Call Finalize

End Sub