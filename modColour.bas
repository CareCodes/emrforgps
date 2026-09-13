Attribute VB_Name = "modColour"
' modRGB - Return the integer color components of a long color variable
' 1996/01/17 Copyright 1996-1998, Larry Rebich
' 1996/07/15 Use Tip 171: Determining RGB Color Values, MSDN July 1996. Larry

    Option Explicit
    DefInt A-Z
    
    Dim miRed As Integer, miGreen As Integer, miBlue As Integer
    Const csComma = ", "
    
Public Function ColorCodeToRGBValue(sColorCodeString As String) As Long
' 96/06/17 Return the RGB value from the color code string
' 96/06/18 Handle even in no space after the comma. Larry.
    Dim i As Integer
    Dim cs As String
    cs = Trim$(csComma)
    Dim sValue As String                'modified from another place in SIMPL
    sValue = sColorCodeString
    i = InStr(sValue, cs)               'get the red value
    miRed = Val(Left$(sValue, i - 1))
    sValue = Mid$(sValue, i + 1)
    i = InStr(sValue, cs)               'get the green value
    miGreen = Val(Left$(sValue, i - 1))
    sValue = Mid$(sValue, i + 1)
    miBlue = Val(sValue)                 'and remainder is blue
    ColorCodeToRGBValue = RGB(miRed, miGreen, miBlue)
End Function
Public Function ColorCodeToRGBString(lColorCode As Long) As String
'  96/06/17 Used to store colors in the registry, return: "128, 128, 255"
    Dim sTemp As String
    
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        sTemp = miRed & csComma & miGreen & csComma & miBlue
        ColorCodeToRGBString = sTemp
    End If
End Function
Public Function ColorCodeToRGB(lColorCode As Long, iRed As Integer, iGreen As Integer, iBlue As Integer) As Boolean
' 96/01/16 Return the individual colors for lColorCode.
' Enter with:
'   lColorCode contains the color to be converted
'
' Return:
'   iRed      contains the red component
'   iGreen             the green component
'   iBlue              the blue component
'
' 96/07/15 Use Tip 171: Determining RGB Color Values, MSDN July 1996. Larry
    Dim lColor As Long          'Tip 171: Determining RGB Color Values, MSDN July 1996
    lColor = lColorCode         'work long
    iRed = lColor Mod &H100     'get red component
    lColor = lColor \ &H100     'divide
    iGreen = lColor Mod &H100   'get green component
    lColor = lColor \ &H100     'divide
    iBlue = lColor Mod &H100    'get blue component
    
    ColorCodeToRGB = True
End Function

Public Function RGBtoColorCode(iRed As Integer, iGreen As Integer, iBlue As Integer) As Long
' Return the long using iRed, iGreen and iBlue, same as the RGB function but added for completness
    On Error Resume Next
    RGBtoColorCode = RGB(iRed, iGreen, iBlue)
End Function

Public Function RGB_Red(lColorCode As Long) As Integer
' Return the red component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Red = miRed
    End If
End Function

Public Function RGB_Green(lColorCode As Long) As Integer
' Return the green component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Green = miGreen
    End If
End Function

Public Function RGB_Blue(lColorCode As Long) As Integer
' Return the blue component of lColorCode
    If ColorCodeToRGB(lColorCode, miRed, miGreen, miBlue) Then
        RGB_Blue = miBlue
    End If
End Function




