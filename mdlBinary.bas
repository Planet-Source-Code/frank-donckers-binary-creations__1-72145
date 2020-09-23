Attribute VB_Name = "mdlBinary"
Option Explicit

'**********************************************************************
'************            MDLBINARY.BAS             ********************
'************                                      ********************
'************  Module with tasks for binary files  ********************
'************  Create binary file                  ********************
'************  Extract files from binary file      ********************
'************  Add files to binary file            ********************
'************  Remove files from binary file       ********************
'**********************************************************************


'**********************************************************************
'Declarations
'**********************************************************************

'Structures
Public Type BinFileStructure    'Structure holds filesize and number of contained files in binary file
    intNumFiles As Integer      'Number of files
    lngFileSize As Long         'Size of the file
End Type

Public Type BinFileData         'Structure holds size, startposition and name for each file contained in the binary file
    lngFileSize As Long         'Size of Chunck of stored data
    lngFileStart As Long        'Start of the Chunk
    strFileName As String * 16  'Filename stored file
End Type

Public Type bytFileData         'Structure to hold the data for each file contained in the binary file
    bytData() As Byte           'Data of each file
End Type

'Numerics
Public i As Integer             'Counter
Public ii As Integer            'Counter
Public iii As Integer           'Counter

'Alfanumerics
Public Const strBinPath As String = "c:\bin"                    'Holds the standard path to the binary files
Public Const strExtractToPath As String = "c:\bin\extracted"    'Holds the standard path to the extraction folder


'//////////////////////////////////////////
'Subroutine Removing files from binary file
'//////////////////////////////////////////
Public Sub RemoveFromBinary(strBinaryFile As String, strToKeep() As String _
    , strDataHolder() As String, tmpFileHead As BinFileStructure, tmpInfoHead() As BinFileData)
    
    'Declarations
    Dim intOpen_Binary_File As Integer  'File number for the binary file
    Dim bytDataToCombine() As bytFileData  'Data in the selected files
    Dim FileHead As BinFileStructure       'Fileheader
    Dim InfoHead() As BinFileData     'Infoheader
    Dim lngFileStart As Long            'Startposition of file
    
    'Resize data array to files already in the binary file + files to be added to the binary file
    ReDim bytDataToCombine(UBound(strToKeep))
    
    'On errors jump to the Error handler
    On Local Error GoTo ErrOut
    
    'Store the data of each file to keep from tmpInfoHead in filehead
    For i = LBound(strToKeep) To UBound(strToKeep)
        For ii = LBound(tmpInfoHead) To UBound(tmpInfoHead)
            'Store only the files listed in strToKeep
            If strToKeep(i) = tmpInfoHead(ii).strFileName Then
                'Resize of the data arrays
                ReDim bytDataToCombine(i).bytData(Len(strDataHolder(ii)))
                'Put the data of the dataholder into the data array for the bynary file
                bytDataToCombine(i).bytData = strDataHolder(ii)
                'Set up the info headers.filesize
                FileHead.lngFileSize = FileHead.lngFileSize + (UBound(bytDataToCombine(i).bytData) + 1)
                Exit For
            End If
        Next ii
    Next i
    
    'Set up the filehead filecount
    FileHead.intNumFiles = UBound(strToKeep) + 1
    
    'Set up the filehead filesize
    FileHead.lngFileSize = FileHead.lngFileSize + (6) + (FileHead.intNumFiles * 24)
    
     'Set up the infohead
    ReDim InfoHead(FileHead.intNumFiles - 1)
    lngFileStart = (6) + (FileHead.intNumFiles * 24) + 1
    
    'Set up the infohead filenames, filestarts and filesizes
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        InfoHead(i).strFileName = strToKeep(i)
        InfoHead(i).lngFileSize = UBound(bytDataToCombine(i).bytData) + 1
        InfoHead(i).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(i).lngFileSize
     Next
    On Error Resume Next
    
    'Delete the original binary file
    Kill strBinaryFile
    
    'Get a free file number for the new binary file
    intOpen_Binary_File = FreeFile
    
    'Open the new binary file
    Open strBinaryFile For Binary Access Write Lock Write As intOpen_Binary_File
        'Store the FileHead in the binary file
        Put intOpen_Binary_File, 1, FileHead
        'Store the InfoHead in the binary file
        Put intOpen_Binary_File, , InfoHead
        'Store the data in the binary file
        For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
            Put intOpen_Binary_File, , bytDataToCombine(i).bytData
        Next
        'Close the binary file
     Close intOpen_Binary_File
    'No errors, Exit this sub
    Exit Sub
    
ErrOut:
    'Display the error message
    strMsg = MsgBox("RemoveFromBinary is unable to remove the files from" & vbCrLf & strBinaryFile, vbOKOnly + vbCritical, "Error")
End Sub

'//////////////////////////////////////////
'Subroutine adding files to binary file
'//////////////////////////////////////////
Public Sub AddToBinary(strBinaryFile As String, str_Path_Array() As String, str_File_Array() As String _
    , strDataHolder() As String, tmpFileHead As BinFileStructure, tmpInfoHead() As BinFileData)
    
    'Declarations
    Dim intOpenfileNumber() As Integer  'File numbers to open selected files
    Dim intFileCount As Integer         'Number of files to be added
    Dim intFileCountAll As Integer      'Number of files files to be added + number of files already in the binary file
    Dim intOpen_Binary_File As Integer  'File number for the binary file
    Dim bytDataToCombine() As bytFileData  'Data in the selected files
    Dim FileHead As BinFileStructure       'Fileheader
    Dim InfoHead() As BinFileData     'Infoheader
    Dim lngFileStart As Long
    
    'Set filecount to number of files listed in array of selected files
    intFileCount = UBound(str_File_Array)
    
    'Resize array of free file numbers for opening files
    ReDim intOpenfileNumber(intFileCount)
    
    'Resize data array to files already in the binary file + files to be added to the binary file
    ReDim bytDataToCombine(intFileCount + tmpFileHead.intNumFiles)
    
    'On errors jump to the Error handler
    On Local Error GoTo ErrOut
    
    'Get free file numbers to open the selected files
    For i = LBound(str_File_Array) To UBound(str_File_Array)
        intOpenfileNumber(i) = FreeFile
        Open str_Path_Array(i) & "\" & str_File_Array(i) For Binary Access Read Lock Write As intOpenfileNumber(i)
    Next i
    
    'Resize of the data arrays
    For i = 0 To intFileCount
        'Resize the data array
        ReDim bytDataToCombine(i).bytData(LOF(intOpenfileNumber(i)))
        'Get the data from the file
        Get intOpenfileNumber(i), 1, bytDataToCombine(i).bytData
    Next
    
    'Close the opened files
    For i = 0 To intFileCount
        'Close and delete the files
        Close intOpenfileNumber(i)
    Next
    
    'Add filecount of files to add to the filecount of files already in the binary file
    intFileCountAll = intFileCount + tmpFileHead.intNumFiles
    
    'Set up the filehead filecount
    FileHead.intNumFiles = (intFileCount + tmpFileHead.intNumFiles) + 1
    
    'Add new files to the data array
    ii = 0
    For i = intFileCount + 1 To intFileCountAll
        ReDim bytDataToCombine(i).bytData(Len(strDataHolder(ii)))
        bytDataToCombine(i).bytData = strDataHolder(ii)
        ii = ii + 1
    Next i
    
    'Set up the filehead filecount
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        FileHead.lngFileSize = FileHead.lngFileSize + (UBound(bytDataToCombine(i).bytData) + 1)
    Next
    
    'Setup filehead filesize
    FileHead.lngFileSize = FileHead.lngFileSize + (6) + (FileHead.intNumFiles * 24)
    
    'Set up the infohead filsizes and filestarts
    ReDim InfoHead(FileHead.intNumFiles - 1)
    lngFileStart = (6) + (FileHead.intNumFiles * 24) + 1
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        InfoHead(i).lngFileSize = UBound(bytDataToCombine(i).bytData) + 1
        InfoHead(i).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(i).lngFileSize
    Next
    
    'Set up the infohead filenames for the new files to be added
    For i = LBound(str_File_Array) To UBound(str_File_Array)
        InfoHead(i).strFileName = str_File_Array(i)
    Next
    
    'Set up the infohead filenames for the files already in the binary file
    ii = 0
    For i = UBound(str_File_Array) + 1 To UBound(bytDataToCombine)
        InfoHead(i).strFileName = tmpInfoHead(ii).strFileName
        ii = ii + 1
    Next
    
    'Get a free file number for the new binary file
    intOpen_Binary_File = FreeFile
    
    'Open the new binary file
    Open strBinaryFile For Binary Access Write Lock Write As intOpen_Binary_File
        'Store the data in the binary file
        Put intOpen_Binary_File, 1, FileHead
        'Store the InfoHead in the file
        Put intOpen_Binary_File, , InfoHead
        'Store the data in the binary file
        For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
            Put intOpen_Binary_File, , bytDataToCombine(i).bytData
        Next
        'Close the binary file
     Close intOpen_Binary_File
     
    'No errors, Exit this sub
    Exit Sub
    
ErrOut:
    'Display the error message
    strMsg = MsgBox("AddToBinary is unable to add the files" & vbCrLf & strBinaryFile, vbOKOnly + vbCritical, "Error")
End Sub

'/////////////////////////////////////////////////
'Suboutine reading/extracting files from binary file
'/////////////////////////////////////////////////
Public Sub ReadAndExtractBinary(strBinaryFile As String, strExtractPath As String, blnReadOnly As Boolean _
    , blnErrors As Boolean, strDataHolder() As String, FileHead As BinFileStructure, InfoHead() As BinFileData)
    'Declarations
    Dim intOpenfileNumber As Integer    'File number to put data
    Dim intOpen_Binary_File As Integer  'File number for the binary file
    Dim BytDataHolder() As Byte         'Array for the data of the opened file
   
    'On errors jump to the Error handler
    blnErrors = False
    On Local Error GoTo ErrOut

    'Open the binary file to read and extract
    intOpen_Binary_File = FreeFile
    Open strBinaryFile For Binary Access Read Lock Write As intOpen_Binary_File
    
        'Extract the filehead
        Get intOpen_Binary_File, 1, FileHead
        
        'Check the file for validity (extacting only possible from files created with this programm)
        If LOF(intOpen_Binary_File) <> FileHead.lngFileSize Then
            strMsg = MsgBox("This is an invalid file format.", vbOKOnly + vbExclamation, "Invalid File")
            Exit Sub
        End If
        
        'Resize the InfoHead array
        ReDim InfoHead(FileHead.intNumFiles - 1)
        
        'Extract the infohead
        Get intOpen_Binary_File, , InfoHead
        
        'Resize the data array
        ReDim strDataHolder(UBound(InfoHead))
        
        'Extract the files or get the filenames from the file
            For i = 0 To UBound(InfoHead)
            
                'Resize the data array
                ReDim BytDataHolder(InfoHead(i).lngFileSize - 1)
                
                'Get the data from the file
                Get intOpen_Binary_File, InfoHead(i).lngFileStart, BytDataHolder
                
                'Open a new binary file and store the data if needed
                If blnReadOnly = False Then
                    'Get free file number
                    intOpenfileNumber = FreeFile
                    'Open the new file and put the data in it
                    Open strExtractPath & "\" & InfoHead(i).strFileName For Binary Access Write Lock Write As intOpenfileNumber
                        Put intOpenfileNumber, 1, BytDataHolder
                    Close intOpenfileNumber
                End If
                'Put the data in the dataholder array
                strDataHolder(i) = BytDataHolder
            Next
            
    'Close the binary file
    Close intOpen_Binary_File
    
    'No errors, Exit this sub
    Exit Sub

ErrOut:
    blnErrors = True
    'Display the error message
    strMsg = MsgBox("ReadExtract is unable to extract from" & vbCrLf & strBinaryFile, vbOKOnly + vbCritical, "Error")
End Sub

'/////////////////////////////////////////////////
'Subroutine combining files to binary file
'/////////////////////////////////////////////////
Public Sub CreateBinary(strBinaryFile As String, str_Path_Array() As String, str_File_Array() As String)
   
    'Declarations
    Dim intOpenfileNumber() As Integer  'File number to put data
    Dim intFileCount As Integer         'Number of files to be added
    Dim intOpen_Binary_File As Integer  'File number for the binary file
    Dim bytDataToCombine() As bytFileData  'Array for the data of the opened file
    Dim FileHead As BinFileStructure       'Fileheader
    Dim InfoHead() As BinFileData     'Infoheader
    Dim lngFileStart As Long            'Startposition of file in binary file
    
    'Set filecount to number of files in file array
    intFileCount = UBound(str_File_Array)
    
    'Resize array of free file numbers for opening files
    ReDim intOpenfileNumber(intFileCount)
    
    'Resize data array
    ReDim bytDataToCombine(intFileCount)
    
    'On errors jump to the Error handler
    On Local Error GoTo ErrOut
    
    'Get free file numbers to use for opening the files
    For i = LBound(str_File_Array) To UBound(str_File_Array)
        intOpenfileNumber(i) = FreeFile
        'Open the file
        Open str_Path_Array(i) & "\" & str_File_Array(i) For Binary Access Read Lock Write As intOpenfileNumber(i)
    Next i
    
    'Resize of the data arrays
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        'Resize the data array
        ReDim bytDataToCombine(i).bytData(LOF(intOpenfileNumber(i)))
        'Get the data from the file
        Get intOpenfileNumber(i), 1, bytDataToCombine(i).bytData
    Next
    
    'Close the files
    For i = 0 To intFileCount
        Close intOpenfileNumber(i)
    Next
    
    'Set up the filehead filecount
    FileHead.intNumFiles = intFileCount + 1
    
    'Set up the filehead filesize
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        FileHead.lngFileSize = FileHead.lngFileSize + (UBound(bytDataToCombine(i).bytData) + 1)
    Next
    FileHead.lngFileSize = FileHead.lngFileSize + (6) + (FileHead.intNumFiles * 24)
    
    'Resize infohead to number of files
    ReDim InfoHead(FileHead.intNumFiles - 1)
    
    'Set up the infohead filenames, filestarts and filesizes
    lngFileStart = (6) + (FileHead.intNumFiles * 24) + 1
    For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
        InfoHead(i).lngFileSize = UBound(bytDataToCombine(i).bytData) + 1
        InfoHead(i).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(i).lngFileSize
    Next
    
    'Fill infofead filnames with the filenames in the array of filenames
    For i = LBound(str_File_Array) To UBound(str_File_Array)
        InfoHead(i).strFileName = str_File_Array(i)
    Next
    
    'Get a free file number for the new binary file
    On Error Resume Next
    
    'kill the new binary file when it already exists
    Kill strBinaryFile
    
    'On errors jump to the Error handler
    On Error GoTo ErrOut
    
    'Get free filenumber for the binary file
    intOpen_Binary_File = FreeFile
    
    'Open the new binary file
    Open strBinaryFile For Binary Access Write Lock Write As intOpen_Binary_File
        'Store the filehead in the file
        Put intOpen_Binary_File, 1, FileHead
        'Store the infohead in the file
        Put intOpen_Binary_File, , InfoHead
        'Store the data in the file
        For i = LBound(bytDataToCombine) To UBound(bytDataToCombine)
            Put intOpen_Binary_File, , bytDataToCombine(i).bytData
        Next
        'Close the binary file
     Close intOpen_Binary_File
     
    'No errors, Exit this sub
    Exit Sub
    
ErrOut:
    'Display the error message
    strMsg = MsgBox("CombinBinary is unable to combine to" & vbCrLf & strBinaryFile, vbOKOnly + vbCritical, "Error")
End Sub
