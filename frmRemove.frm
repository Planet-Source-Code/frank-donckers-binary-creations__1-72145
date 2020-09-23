VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRemove 
   Caption         =   "Remove from binary file"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   Icon            =   "frmRemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenBin 
      Caption         =   "Open bin"
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Files to keep"
      Height          =   4215
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   3375
      Begin VB.ListBox lstInBin 
         Height          =   3765
         ItemData        =   "frmRemove.frx":08CA
         Left            =   240
         List            =   "frmRemove.frx":08CC
         TabIndex        =   11
         Top             =   240
         Width           =   2850
      End
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   6855
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
      Left            =   8280
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "> Add File To List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3960
      TabIndex        =   6
      Top             =   1200
      Width           =   2220
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "<< Remove All Files From List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3960
      TabIndex        =   5
      Top             =   3360
      Width           =   2220
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< Remove File From List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   2220
   End
   Begin VB.CommandButton cmdRemoveFromBin 
      Caption         =   "Remove and save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files to remove"
      Height          =   4215
      Left            =   6480
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   3570
         ItemData        =   "frmRemove.frx":08CE
         Left            =   240
         List            =   "frmRemove.frx":08D0
         TabIndex        =   1
         Top             =   360
         Width           =   2850
      End
   End
   Begin MSComDlg.CommonDialog comdiag1 
      Left            =   0
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Binary file"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
'************            FRMADD.FRM                ********************
'************                                      ********************
'************     Adding files to a binary file    ********************
'************                                      ********************
'**********************************************************************

'////////////////////////////////////////////////////////////
'Declarations
'////////////////////////////////////////////////////////////

Private strDataHolder() As String   'Array that holds the data for each file to be added to transfer to mdlbinary
Private FileHead As BinFileStructure   'Fileheader to transfer data to mdlbinary
Private InfoHead() As BinFileData 'Infoheader to transfer data to mdlbinary


'////////////////////////////////////////////////////////////
'Button to open a binary file
'////////////////////////////////////////////////////////////
Private Sub cmdOpenBin_Click()
    Dim blnErrors As Boolean    'Errors found?
    On Error GoTo ErrOut
    
    'Trim the textbox
    txtFile.Text = Trim$(txtFile.Text)
    
    'Check binary file
    If txtFile.Text = "" Or Dir(txtFile.Text, vbNormal) = "" Then
        strMsg = MsgBox("The file" & vbCrLf & txtFile & vbCrLf & "Does not exist", vbOKOnly + vbExclamation, "ErrorMessage")
        txtFile.SetFocus
        Exit Sub
    End If
    
     'Call the read/extraction routine with readonly (no extraction)
    ReadAndExtractBinary txtFile.Text, "", True, blnErrors, strDataHolder(), FileHead, InfoHead()
    
    'Set list
    If blnErrors = True Then Exit Sub
    lstInBin.Clear
    For i = LBound(InfoHead) To UBound(InfoHead)
        lstInBin.AddItem InfoHead(i).strFileName
    Next i
    'Activate apropryat buttons
    cmdAdd.Enabled = True

'Errorhandler
ErrOut:
End Sub

'////////////////////////////////////////////////////////////
'Button to add one file to the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdAdd_Click()
    AddFileToList
End Sub


'////////////////////////////////////////////////////////////
'Button to select the binary file
'////////////////////////////////////////////////////////////
Private Sub cmdBrowseBin_Click()
    On Error GoTo ErrOut
    Dim blnErrors As Boolean 'Errors found?
    comdiag1.ShowOpen
    If comdiag1.FileName <> "" Then
        txtFile = comdiag1.FileName
        
        'Trim the textbox
        txtFile.Text = Trim$(txtFile.Text)
        
        'Check binary file
        If txtFile.Text = "" Or Dir(txtFile.Text, vbNormal) = "" Then
            strMsg = MsgBox("The file" & vbCrLf & txtFile & vbCrLf & "Does not exist", vbOKOnly + vbExclamation, "ErrorMessage")
            txtFile.SetFocus
            Exit Sub
        End If
         
         'Call the read/extraction routine with readonly (no extraction)
        ReadAndExtractBinary txtFile.Text, "", True, blnErrors, strDataHolder(), FileHead, InfoHead()
        
        'Set list
        If blnErrors = True Then Exit Sub
        lstInBin.Clear
        For i = LBound(InfoHead) To UBound(InfoHead)
            lstInBin.AddItem InfoHead(i).strFileName
        Next i
        'Activate apropryat buttons
        cmdAdd.Enabled = True
        
    End If

'Errorhandler
ErrOut:
End Sub

'////////////////////////////////////////////////////////////
'Button to Cancel and return to main form
'////////////////////////////////////////////////////////////
Private Sub cmdCancel_Click()
    frmBinary.Show
    Me.Hide
End Sub

'////////////////////////////////////////////////////////////
'Button to remove all files from the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdClear_Click()
    'Clear the list of files
    For i = 0 To List1.ListCount - 1
        List1.ListIndex = i
        lstInBin.AddItem List1.Text
    Next i
    List1.Clear
    'Activate apropryat buttons
    cmdRemoveFromBin.Enabled = False
End Sub

'////////////////////////////////////////////////////////////
'Button to remove one file from the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
    On Error Resume Next
    If List1.SelCount < 1 Then Exit Sub
    lstInBin.AddItem List1.Text
    List1.RemoveItem (List1.ListIndex)
    List1.Refresh
    If List1.ListCount = 0 Then cmdRemoveFromBin.Enabled = False
End Sub
Private Sub List1_DblClick()
    cmdRemove_Click
End Sub


'//////////////////////////////////////////////////////////////
'Button to start the removing of the files from the binary file
'//////////////////////////////////////////////////////////////
Private Sub cmdRemoveFromBin_Click()
    Dim strToKeep() As String
    'Reset size array of files to keep
    ReDim strToKeep(lstInBin.ListCount - 1)
    'Store the filenames from the list in the array of files to keep
    For i = 0 To lstInBin.ListCount - 1
        lstInBin.ListIndex = i
        strToKeep(i) = lstInBin.Text
    Next i
    
    'Start removing of the files
    RemoveFromBinary txtFile, strToKeep(), strDataHolder(), FileHead, InfoHead()
    
    'Clear lists
    List1.Clear
    lstInBin.Clear
    
    'Set buttons
    cmdRemove.Enabled = False
End Sub

'////////////////////////////////////////////////////////////
'Activation of form
'////////////////////////////////////////////////////////////
Private Sub Form_Activate()
    lstInBin.Clear
    List1.Clear
    'Activate apropryat buttons
    cmdRemove.Enabled = False
    cmdAdd.Enabled = False
    cmdClear.Enabled = False
End Sub

'////////////////////////////////////////////////////////////
'Select file from list to add to filelist
'////////////////////////////////////////////////////////////
Private Sub lstInBin_DblClick()
    AddFileToList
End Sub

'////////////////////////////////////////////////////////////
'Unloading of the form
'////////////////////////////////////////////////////////////
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

'////////////////////////////////////////////////////////////
'Sub adding one file to the filelist
'////////////////////////////////////////////////////////////
Private Sub AddFileToList()
    If lstInBin.SelCount < 1 Then Exit Sub
    If lstInBin.ListCount < 2 Then
            strMsg = MsgBox("Binary file can not be empty", vbOKOnly + vbExclamation, "Attention")
            Exit Sub
    End If
    'If the selected file is already in the filelist do not add
    If List1.ListCount > 0 Then
        For i = 0 To List1.ListCount
            If List1.List(i) = lstInBin.Text Then Exit Sub
        Next i
    End If
    List1.AddItem lstInBin.Text
    lstInBin.RemoveItem (lstInBin.ListIndex)
    'Activate apropryat buttons
    cmdRemoveFromBin.Enabled = True
    cmdClear.Enabled = True
    cmdRemove.Enabled = True
End Sub

