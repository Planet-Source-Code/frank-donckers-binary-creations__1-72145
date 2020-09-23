VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdd 
   Caption         =   "Add to binary file"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenBin 
      Caption         =   "Open bin"
      Height          =   375
      Left            =   9600
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Files in binary file"
      Height          =   2895
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   10335
      Begin VB.ListBox lstInBin 
         Enabled         =   0   'False
         Height          =   2400
         ItemData        =   "frmAdd.frx":08CA
         Left            =   240
         List            =   "frmAdd.frx":08CC
         TabIndex        =   16
         Top             =   240
         Width           =   9810
      End
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   120
      Width           =   7695
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
      Left            =   9000
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "> Add File To List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4800
      TabIndex        =   11
      Top             =   4800
      Width           =   2220
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "<< Remove All Files From List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4800
      TabIndex        =   10
      Top             =   6960
      Width           =   2220
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "< Remove File From List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4800
      TabIndex        =   9
      Top             =   6240
      Width           =   2220
   End
   Begin VB.CommandButton cmdAddAll 
      Caption         =   ">> Add All Files To List"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4800
      TabIndex        =   8
      Top             =   5520
      Width           =   2220
   End
   Begin VB.CommandButton cmdAddtoBin 
      Caption         =   "Add and save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files to add"
      Height          =   4215
      Left            =   7200
      TabIndex        =   4
      Top             =   3840
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   3570
         ItemData        =   "frmAdd.frx":08CE
         Left            =   240
         List            =   "frmAdd.frx":08D0
         TabIndex        =   5
         Top             =   360
         Width           =   2850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files to select"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   4095
      Begin VB.DirListBox Dir1 
         Height          =   3240
         Left            =   240
         TabIndex        =   3
         Top             =   735
         Width           =   1905
      End
      Begin VB.FileListBox File1 
         Height          =   3600
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1620
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1920
      End
   End
   Begin MSComDlg.CommonDialog comdiag1 
      Left            =   0
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Binary file"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdd"
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

Private strDataHolder() As String       'Array that holds the data for each file to be added to transfer to mdlbinary
Private FileHead As BinFileStructure    'Fileheader to transfer data to mdlbinary
Private InfoHead() As BinFileData       'Infoheader to transfer data to mdlbinary


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
    cmdAdd.Enabled = True
    cmdAddAll.Enabled = True
    
'Errorhandler
ErrOut:
End Sub

'////////////////////////////////////////////////////////////
'Button to add one file to the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdAdd_Click()
    Call AddFileToList
End Sub

'////////////////////////////////////////////////////////////
'Button to add all files to the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdAddAll_Click()
    Call AddAllFilesToList
End Sub

'////////////////////////////////////////////////////////////
'Button to add all files from the filelist to the binary file
'////////////////////////////////////////////////////////////
Private Sub cmdAddtoBin_Click()
    
    'Start adding the files
    AddToBinary txtFile, Files.filePathArray, Files.fileNameArray, strDataHolder(), FileHead, InfoHead()
    
    'Set buttons
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
    cmdAddtoBin.Enabled = False
    
    'Clear lists
    List1.Clear
    lstInBin.Clear
    txtFile = ""
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
    End If
    
    'Set buttons
    cmdAdd.Enabled = True
    cmdAddAll.Enabled = True
    
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
    List1.Clear
    
    'Clear the file arrays
    ReDim Files.filePathArray(0)
    ReDim Files.fileNameArray(0)
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
    cmdAddtoBin.Enabled = False
End Sub

'////////////////////////////////////////////////////////////
'Button to remove one file from the filelist
'////////////////////////////////////////////////////////////
Private Sub cmdRemove_Click()
    On Error Resume Next
TryAgain:
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            'Remove the file from the list
            List1.RemoveItem (i)
            
            'Remove the file from the arrays by using a temporary arrays
            'Resize the temporary arrays
            ReDim TempFiles.fileNameArray(UBound(Files.fileNameArray))
            ReDim TempFiles.filePathArray(UBound(Files.filePathArray))
            
            'move filenames to temporary array
            For ii = LBound(Files.fileNameArray) To UBound(Files.fileNameArray)
                TempFiles.fileNameArray(ii) = Files.fileNameArray(ii)
                TempFiles.filePathArray(ii) = Files.filePathArray(ii)
            Next
            Dim CurrNum
            Dim CurrNum2
            CurrNum = 0
            CurrNum2 = 0
            
            'Move the files back from the temporary arrays to the file arrays
            For ii = LBound(Files.fileNameArray) To UBound(Files.fileNameArray)
                 If ii <> i Then
                    Files.fileNameArray(CurrNum2) = TempFiles.fileNameArray(CurrNum)
                    Files.filePathArray(CurrNum2) = TempFiles.filePathArray(CurrNum)
                    CurrNum2 = CurrNum2 + 1
                 End If
                 CurrNum = CurrNum + 1
            Next
            
            'Resize the file arrays
            ReDim Preserve Files.filePathArray(UBound(Files.fileNameArray) - 1)
            ReDim Preserve Files.fileNameArray(UBound(Files.fileNameArray) - 1)
            GoTo TryAgain
        End If
    Next i
    List1.Refresh
    If List1.ListCount = 0 Then
        cmdRemove.Enabled = False
        cmdClear.Enabled = False
        cmdAddtoBin.Enabled = False
    End If
    
End Sub
Private Sub List1_DblClick()
    cmdRemove_Click
End Sub

'////////////////////////////////////////////////////////////
'Changing directorylist
'////////////////////////////////////////////////////////////
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

'////////////////////////////////////////////////////////////
'Changing drivebox
'////////////////////////////////////////////////////////////
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

'////////////////////////////////////////////////////////////
'Select file from list to add to filelist
'////////////////////////////////////////////////////////////
Private Sub File1_DblClick()
    AddFileToList
End Sub

'////////////////////////////////////////////////////////////
'Sub adding one file to the filelist
'and adding to the path/file arrays
'////////////////////////////////////////////////////////////
Private Sub AddFileToList()
    'If the selected file is already in the filelist do not add
    If List1.ListCount > 0 Then
        For i = LBound(Files.fileNameArray) To UBound(Files.fileNameArray)
            If Files.fileNameArray(i) = "\" & File1.FileName Then Exit Sub
        Next i
    End If
    
    'Add the filename to the list
    List1.AddItem File1.FileName
    
    'Resize the filename array
    ReDim Preserve Files.filePathArray(List1.ListCount - 1)
    
    'Resize the pathname array
    ReDim Preserve Files.fileNameArray(List1.ListCount - 1)
    
    'Add file to the arrays
    Files.fileNameArray(List1.ListCount - 1) = "\" & File1.FileName
    Files.filePathArray(List1.ListCount - 1) = File1.Path
    If List1.ListCount > 0 Then
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        cmdAddtoBin.Enabled = True
    End If

End Sub

'////////////////////////////////////////////////////////////
'Sub adding multiple files to the filelist
'and adding to the path/file arrays
'////////////////////////////////////////////////////////////
Private Sub AddAllFilesToList()
    On Error Resume Next
    For i = 0 To File1.ListCount - 1
        File1.Selected(i) = True
        For ii = LBound(Files.fileNameArray) To UBound(Files.fileNameArray)
            If Files.fileNameArray(i) <> "\" & File1.FileName Then
                'Add the filename
                List1.AddItem File1.FileName
                
                'Resize the filename array
                ReDim Preserve Files.filePathArray(List1.ListCount - 1)
                
                'Resize the path array
                ReDim Preserve Files.fileNameArray(List1.ListCount - 1)
                
                'Add file to the arrays
                Files.fileNameArray(List1.ListCount - 1) = "\" & File1.FileName
                Files.filePathArray(List1.ListCount - 1) = File1.Path
                Exit For
            End If
        Next ii
    Next i
    If List1.ListCount > 0 Then
        cmdRemove.Enabled = True
        cmdClear.Enabled = True
        cmdAddtoBin.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    List1.Clear
    lstInBin.Clear
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
    cmdAdd.Enabled = False
    cmdAddAll.Enabled = False
    cmdAddtoBin.Enabled = False
End Sub

'////////////////////////////////////////////////////////////
'Unloading of the form
'////////////////////////////////////////////////////////////
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

