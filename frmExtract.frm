VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExtract 
   Caption         =   "Extract/read from binary file"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   Icon            =   "frmExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExtract 
      Caption         =   "Read only don 't extract"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   1230
      ItemData        =   "frmExtract.frx":08CA
      Left            =   1680
      List            =   "frmExtract.frx":08CC
      TabIndex        =   8
      Top             =   1560
      Width           =   7935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowsePath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   7455
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract/read files"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowseBin 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin MSComDlg.CommonDialog comdiag1 
      Left            =   0
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "Extract to path"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Binary file"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****************************************************************************
'************               FRMEXTRACT.FRM                ********************
'************                                             ********************
'************ Form to extract/read files from binary file ********************
'************                                             ********************
'*****************************************************************************

'/////////////////////////////////////////////////////////////
'Declarations
'/////////////////////////////////////////////////////////////

Private strDataHolder() As String   'Array that holds the data for each file to be added to transfer to mdlbinary
Private FileHead As BinFileStructure   'Fileheader to transfer data to mdlbinary
Private InfoHead() As BinFileData 'Infoheader to transfer data to mdlbinary

'////////////////////////////////////////////////////////////
'Button to select the binary file
'////////////////////////////////////////////////////////////
Private Sub cmdBrowseBin_Click()
    On Error GoTo ErrOut
    comdiag1.ShowOpen
    If comdiag1.FileName <> "" Then txtFile = comdiag1.FileName

'Errorhandler
ErrOut:
End Sub

'////////////////////////////////////////////////////////////
'Button to select the folder to extract the files to
'////////////////////////////////////////////////////////////
Private Sub cmdBrowsePath_Click()
    If Dir(Trim$(txtPath), vbDirectory) <> "" Then frmBrowse.Dir1.Path = Trim$(txtPath)
    frmBrowse.Dir1.Refresh
    frmBrowse.Show 1, Me
    
    'Extract all the files to the directory that the binary file is in.
    If OKCancel = True Then txtPath = strBrowsePath
End Sub

'////////////////////////////////////////////////////////////
'Cancel and return to main form
'////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    frmBinary.Show
    Me.Hide
End Sub

'////////////////////////////////////////////////////////////
'Unloading of the form
'////////////////////////////////////////////////////////////

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

'////////////////////////////////////////////////////////////
'Button to start extracting
'////////////////////////////////////////////////////////////
Private Sub cmdExtract_Click()
    Dim blnReadOnly As Boolean  'Read/extract?
    Dim blnErrors As Boolean    'Errors found?
    
    'Trim the textboxes
    txtPath.Text = Trim$(txtPath.Text)
    txtFile.Text = Trim$(txtFile.Text)
    
    'Check binary file
    If txtFile.Text = "" Or Dir(txtFile.Text, vbNormal) = "" Then
        strMsg = MsgBox("The file" & vbCrLf & txtFile & vbCrLf & "Does not exist", vbOKOnly + vbExclamation, "ErrorMessage")
        txtFile.SetFocus
        Exit Sub
    End If
    
    'Readonly or extract?
    If chkExtract.Value = Checked Then blnReadOnly = True
    
    'Check path to extract to
    If blnReadOnly = False Then
        If txtPath.Text = "" Or Dir(txtPath.Text, vbDirectory) = "" Then
            strMsg = MsgBox("The path" & vbCrLf & txtPath & vbCrLf & "Does not exist", vbOKOnly + vbExclamation, "ErrorMessage")
            txtPath.SetFocus
            Exit Sub
        End If
    End If
    lblFiles.Caption = ""
    
    Start reading / extracting
    ReadAndExtractBinary txtFile.Text, txtPath, blnReadOnly, blnErrors, strDataHolder(), FileHead, InfoHead()

    If blnErrors = True Then Exit Sub
    
    'Set list
    lblFiles.Caption = "Files in " & txtFile.Text & " :"
    If blnReadOnly = False Then lblFiles.Caption = "Files extracted to " & txtPath.Text & " :"
    lstFiles.Clear
    On Error Resume Next
    For i = LBound(InfoHead) To UBound(InfoHead)
        lstFiles.AddItem InfoHead(i).strFileName
    Next i
End Sub

