VERSION 5.00
Begin VB.Form frmBinary 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary Creations"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBinary.frx":08CA
   ScaleHeight     =   3690
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove files from binary"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add files to binary"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract/read files from binary"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdCombine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create new binary"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'************          FRMBINARY.FRM               ********************
'************                                      ********************
'************     Main form of the application     ********************
'************                                      ********************
'**********************************************************************

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub

'////////////////////////////////////////////////////////////
'Declarations
'////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////
'Button to open the Add to binary form
'////////////////////////////////////////////////////////////
Private Sub cmdAdd_Click()
    Me.Hide
    frmAdd.Show
    Exit Sub
End Sub

'////////////////////////////////////////////////////////////
'Button to open the make new binary form
'////////////////////////////////////////////////////////////
Private Sub cmdCombine_Click()
    Me.Hide
    frmCreate.Show
End Sub

'////////////////////////////////////////////////////////////
'Button to exit the application
'////////////////////////////////////////////////////////////
Private Sub cmdExit_Click()
    End
End Sub

'////////////////////////////////////////////////////////////
'Button to open the axtract from binary form
'////////////////////////////////////////////////////////////
Private Sub cmdExtract_Click()
    Me.Hide
    frmExtract.Show
End Sub

'////////////////////////////////////////////////////////////
'Button to open the remove from binary form
'////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
    Me.Hide
    frmRemove.Show
End Sub

'////////////////////////////////////////////////////////////
'Loading of the form
'////////////////////////////////////////////////////////////
Private Sub Form_Load()
    frmExtract.txtPath = App.Path & "\extracted"
End Sub

'////////////////////////////////////////////////////////////
'Unloading of the form
'////////////////////////////////////////////////////////////
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
