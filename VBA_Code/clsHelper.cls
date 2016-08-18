VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ===========================================================================================
'
' Author : Christophe Avonture
'
' Date   : 17 August 2016
'
' Aim    : Helper class i.e. contains a lot of functions and subroutines that will be used
'          by the macro
'
' ===========================================================================================

Option Compare Text
Option Base 1
Option Explicit

' ----------------------------------------------------------------------
'
' Return Windows special folder like f.i. "MyDocuments"
'
' Special folders are : AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup, Desktop, Favorites
'   Fonts, MyDocuments, NetHood, PrintHood, Programs, Recent, SendTo, StartMenu, Startup, Templates
'
' IMPORTANT : PARAMETER SHOULD BE A VARIANT; NOT A STRING
'
' ----------------------------------------------------------------------

Public Function GetWindowsSpecialFolders(Optional ByVal sFolder As Variant = "MyDocuments") As String

Dim oWSH As Object      ' Requires Windows Script Host Object Model
Dim sReturn As String

   Set oWSH = CreateObject("Wscript.Shell")
   sReturn = oWSH.SpecialFolders(sFolder) & "\"
   Set oWSH = Nothing
   
   GetWindowsSpecialFolders = sReturn

End Function

' ------------------------------------------------------------------------------
'
' Check if a folder exists and return true in that case
'
' ------------------------------------------------------------------------------

Public Function FolderExists(ByVal sFolderName As String) As Boolean

Dim objFSO As Object

   Set objFSO = CreateObject("Scripting.FileSystemObject")
     
   FolderExists = objFSO.FolderExists(sFolderName)
   
   Set objFSO = Nothing
   
End Function

' ------------------------------------------------------------------------------
'
' Check if a file exists and return true in that case
'
' ------------------------------------------------------------------------------

Public Function FileExists(ByVal sFileName As String) As Boolean

Dim objFSO As Object

   Set objFSO = CreateObject("Scripting.FileSystemObject")
     
   FileExists = objFSO.FileExists(sFileName)
   
   Set objFSO = Nothing
   
End Function

' ------------------------------------------------------------------------------
'
' Recursive folder's creation.
'
' ------------------------------------------------------------------------------

Public Function MakeFolder(FolderName As String) As Boolean
    
Dim objFSO As New FileSystemObject  ' <== Requires reference to Microsoft Scripting runtime
Dim bResult As Boolean

   bResult = False
   
   If Not objFSO.FolderExists(FolderName) Then
      If MakeFolder(objFSO.GetParentFolderName(FolderName)) Then
         bResult = True
         Call objFSO.CreateFolder(FolderName)
      End If
   Else
      bResult = True
   End If
   
   MakeFolder = True

End Function

' ------------------------------------------------------------------------------
'
' Read a text file and return its content
'
' Sample code
' -----------
'
' Dim sFileContent As String
'
'    sFileContent = cHelper.GetFileContent("c:\temp\readme.txt")
'
' ------------------------------------------------------------------------------

Public Function GetFileContent(ByVal sFileName As String) As String

Dim fso As Object
Dim ts As Object

   If (Me.FileExists(sFileName)) Then

      Set fso = CreateObject("Scripting.FileSystemObject")
      Set ts = fso.GetFile(sFileName).OpenAsTextStream(1, -2)
   
      GetFileContent = ts.ReadAll
      ts.Close
      
      Set ts = Nothing
      Set fso = Nothing
   
   Else
   
      GetFileContent = vbNullString
      
   End If
   
End Function