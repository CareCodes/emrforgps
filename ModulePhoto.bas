Attribute VB_Name = "ModulePhoto"
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const MAX_PATH = 260
Public Const BLOCK_SIZE = 10000
'------------------------------------------------------------
' Image Retriving and filling functions
' Function Name : FillPhoto
'   Use : Retrive Stored Photos From Access Database
'   Stores Photos : In Temporary File
'   Parameters:
'   rstMain as Recordset : The Recordset that contains Photos
'   PFName as String : Name of the Field that has the Photo
'   SizeField : Name of the Field That contains size of Photo
'============================================================
' Sub Function Name: GetPhoto
'   Use : Store Photo from a file to Access database
'   Parameters:
'   fileName as String : Name of the File that contains Photo
'   rstMain as Recordset : recordset that has to be updated
'   FieldName as string : Name of the Image Field
'   SizeField :Name of the Size field that contains photo size
'-------------------------------------------------------------

Public Sub FillPhoto(rstMain As Recordset, PFName As String, SizeField As String, picEmp As PictureBox)
On Error GoTo Handler
Dim bytes() As Byte
Dim file_name As String
Dim file_num As Integer
Dim file_length As Long
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long
Dim hgt As Single

    'me.imgPhoto.Visible = False
'    Screen.MousePointer = vbHourglass
'    DoEvents

    ' Get a temporary file name.
    file_name = TemporaryFileName()

    ' Open the file.
    file_num = FreeFile
    Open file_name For Binary As #file_num

    ' Copy the data into the file.
    file_length = rstMain(SizeField)
    num_blocks = file_length / BLOCK_SIZE
    left_over = file_length Mod BLOCK_SIZE

    For block_num = 1 To num_blocks
        bytes() = rstMain(PFName).GetChunk(BLOCK_SIZE)
        Put #file_num, , bytes()
    Next block_num

    If left_over > 0 Then
        bytes() = rstMain(PFName).GetChunk(left_over)
        Put #file_num, , bytes()
    End If

    Close #file_num

    picEmp.Picture = LoadPicture(file_name)
 
    Screen.MousePointer = vbDefault
Exit Sub

Handler:
    Debug.Print Err.Description
    Resume Next
End Sub

Public Sub GetPhoto(filename As String, rstMain As Recordset, FieldName As String, SizeField As String)
On Error GoTo Handler
Dim file_num As String
Dim file_length As Long
Dim bytes() As Byte
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long

    file_num = FreeFile
    Open filename For Binary Access Read As #file_num

    file_length = LOF(file_num)
    If file_length > 0 Then
        num_blocks = file_length / BLOCK_SIZE
        left_over = file_length Mod BLOCK_SIZE

        rstMain(SizeField) = file_length

        ReDim bytes(BLOCK_SIZE)
        For block_num = 1 To num_blocks
            Get #file_num, , bytes()
            rstMain(FieldName).AppendChunk bytes()
        Next block_num

        If left_over > 0 Then
            ReDim bytes(left_over)
            Get #file_num, , bytes()
            rstMain(FieldName).AppendChunk bytes()
        End If

        'rstEmployee.Update
        Close #file_num
    End If
Exit Sub

Handler:
    MsgBox Err.Description
    Resume
   Debug.Print Err.Description

End Sub

Public Function TemporaryFileName() As String
Dim temp_path As String
Dim temp_file As String
Dim length As Long

    ' Get the temporary file path.
    temp_path = VBA.Space$(MAX_PATH)
    length = GetTempPath(MAX_PATH, temp_path)
    temp_path = Left$(temp_path, length)

    ' Get the file name.
    temp_file = VBA.Space$(MAX_PATH)
    GetTempFileName temp_path, "per", 0, temp_file
    TemporaryFileName = Left$(temp_file, InStr(temp_file, VBA.Chr$(0)) - 1)
End Function

Public Sub UnRGB(ByVal color As OLE_COLOR, ByRef R As Integer, ByRef G As Integer, ByRef b As Integer)
   b = color \ 65536
   G = (color \ 256) Mod 256
   R = color Mod 256
End Sub

Public Sub SetBrightness(ByVal pic1 As PictureBox, ByVal pic2 As PictureBox, ByVal brightness As Single)
Dim fraction As Single
Dim X As Integer
Dim Y As Integer
Dim R As Integer
Dim G As Integer
Dim b As Integer

    If brightness < 0 Then
        ' Darken.
        fraction = (100 + brightness) / 100
        For Y = 0 To pic1.ScaleHeight - 1
            For X = 0 To pic1.ScaleWidth - 1
                'DoEvents
                UnRGB pic1.Point(X, Y), R, G, b
                R = R * fraction
                G = G * fraction
                b = b * fraction
                pic2.PSet (X, Y), RGB(R, G, b)
            Next X
        Next Y
    Else
        ' Brighten.
        fraction = brightness / 100
        For Y = 0 To pic1.ScaleHeight - 1
            For X = 0 To pic1.ScaleWidth - 1
                'DoEvents
                UnRGB pic1.Point(X, Y), R, G, b
                R = R + (255 - R) * fraction
                G = G + (255 - G) * fraction
                b = b + (255 - b) * fraction
                pic2.PSet (X, Y), RGB(R, G, b)
            Next X
        Next Y
    End If
    pic2.Picture = pic2.Image
End Sub


















Public Sub FillImage(rstMain As Recordset, PFName As String, SizeField As String, picEmp As Image)
On Error GoTo Handler
Dim bytes() As Byte
Dim file_name As String
Dim file_num As Integer
Dim file_length As Long
Dim num_blocks As Long
Dim left_over As Long
Dim block_num As Long
Dim hgt As Single

    'me.imgPhoto.Visible = False
'    Screen.MousePointer = vbHourglass
'    DoEvents

    ' Get a temporary file name.
    file_name = TemporaryFileName()

    ' Open the file.
    file_num = FreeFile
    Open file_name For Binary As #file_num

    ' Copy the data into the file.
    file_length = rstMain(SizeField)
    num_blocks = file_length / BLOCK_SIZE
    left_over = file_length Mod BLOCK_SIZE

    For block_num = 1 To num_blocks
        bytes() = rstMain(PFName).GetChunk(BLOCK_SIZE)
        Put #file_num, , bytes()
    Next block_num

    If left_over > 0 Then
        bytes() = rstMain(PFName).GetChunk(left_over)
        Put #file_num, , bytes()
    End If

    Close #file_num

    picEmp.Picture = LoadPicture(file_name)
 
    Screen.MousePointer = vbDefault
Exit Sub

Handler:
    Debug.Print Err.Description
    Resume Next
End Sub

