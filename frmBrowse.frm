VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Browse for folder"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5115
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
'************           FRMBROWSE.FRM              ********************
'************                                      ********************
'************       Form to select a folder        ********************
'************                                      ********************
'**********************************************************************

'/////////////////////////////////////////////////////////////
'Declarations
'/////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////
'Button to cancel and return to main form
'/////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    OKCancel = False
    Unload Me
End Sub

'/////////////////////////////////////////////////////////////
'Button to confirm the selected folder and return to main form
'/////////////////////////////////////////////////////////////
Private Sub cmdOK_Click()
    OKCancel = True
    strBrowsePath = Dir1.Path
    Unload Me
End Sub

'/////////////////////////////////////////////////////////////
'Changing directorylist
'/////////////////////////////////////////////////////////////
Private Sub Dir1_Change()
    cmdOK.Enabled = True
End Sub

'/////////////////////////////////////////////////////////////
'Changing drivebox
'/////////////////////////////////////////////////////////////
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

'/////////////////////////////////////////////////////////////
'Loading of the form
'/////////////////////////////////////////////////////////////
Private Sub Form_Load()
    Dir1.Path = "c:\"
End Sub
