Attribute VB_Name = "mdlMain"
Option Explicit

'*********************************************************************
'************            MDLMAIN.BAS              ********************
'************                                     ********************
'************   Module with main declarations     ********************
'************                                     ********************
'*********************************************************************

'*********************************************************************
'Declarations
'*********************************************************************

Public OKCancel As Boolean      'Boolean to check if Cancel button is selected in frmBrowse (folder browser)
Public strBrowsePath As String  'Folder selected in frmBrowse (folder browser)

Public Type FileInfo            'Structure FyleInfo holds arrays of filenames/paths for selected files
    fileNameArray() As String   'Array file names of selected files
    filePathArray() As String   'Array File paths of selected files
End Type

Public Files As FileInfo        'Array of selected files to ad to binary file
Public TempFiles As FileInfo    'Temporary array to remove files from the array selected files

Public strMsg As String         'Messages
