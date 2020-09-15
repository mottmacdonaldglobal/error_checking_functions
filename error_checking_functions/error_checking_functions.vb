Attribute VB_Name = "GetDocumentInfo"
'FUNCTIONS FOR EASILY GETTING DOCUMENT INFORMATION

'==================================================================================
'GetFilename Function
'----------------------------------------------------------------------------------
'Purpose:   Returns the filename of the document for use in headings etc
'
'Author:    Rachel J Arthur
'
'Notes:     -
'
'----------------------------------------------------------------------------------
'Parameters
'----------------------------------------------------------------------------------
'
'blnGetExtension:       If TRUE, returns filename with extension;           Optional Boolean
'                           if FALSE, returns it without extension
'
'----------------------------------------------------------------------------------
'Returns
'----------------------------------------------------------------------------------
'
'Returns:               Filename in which the function is included          String
'
'----------------------------------------------------------------------------------
'Revision History
'----------------------------------------------------------------------------------
'
'Version 1.0.0      01/05/2019      RJA     Initial release
'
'----------------------------------------------------------------------------------
Public Function GetFilename(Optional blnGetExtension As Boolean = True) As String

    If blnGetExtension Then
        GetFilename = ThisWorkbook.Name
    Else
        MyFilename = ThisWorkbook.Name
        GetFilename = Left(MyFilename, InStr(1, MyFilename, ".", vbTextCompare) - 1)
    End If
   
End Function

'==================================================================================
'GetTabname Function
'----------------------------------------------------------------------------------
'Purpose:   Returns the tab or sheet name for use in headings etc
'
'Author:    Rachel J Arthur
'
'Notes:     -
'
'----------------------------------------------------------------------------------
'Parameters
'----------------------------------------------------------------------------------
'
'intTabIndex:   Sheet number of the tab name that you want,                 Optional Integer
'                   if not the tab on which the cell is located
'
'----------------------------------------------------------------------------------
'Returns
'----------------------------------------------------------------------------------
'
'Returns:       Tab or sheet name                                           String
'
'----------------------------------------------------------------------------------
'Revision History
'----------------------------------------------------------------------------------
'
'Version 1.0.0      01/05/2019      RJA     Initial release
'
'----------------------------------------------------------------------------------
Public Function GetTabname(Optional intTabIndex As Integer = 0) As String
    
    'If no index number is supplied, returns the name of the sheet in which the function is included
    If intTabIndex = 0 Then
        GetTabname = Application.Caller.Worksheet.Name
    Else
    'Otherwise returns the tabname of the index
        GetTabname = ThisWorkbook.Sheets(intTabIndex).Name
    End If
    
End Function


'==================================================================================
'GetSheetNumber Function
'----------------------------------------------------------------------------------
'Purpose:   Returns the index number of the sheet for information
'
'Author:    Rachel J Arthur
'
'Notes:     -
'
'----------------------------------------------------------------------------------
'Parameters
'----------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------
'Returns
'----------------------------------------------------------------------------------
'
'Returns:       Tab or sheet number                                         Integer
'
'----------------------------------------------------------------------------------
'Revision History
'----------------------------------------------------------------------------------
'
'Version 1.0.0      01/05/2019      RJA     Initial release
'
'----------------------------------------------------------------------------------

Public Function GetSheetNumber() As Integer
    GetSheetNumber = ThisWorksheet().Index
End Function
