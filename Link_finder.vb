'Finds Linked Objects
Sub Link_Finder()
Dim oShape As InlineShape, strMsg As String
Debug.Print Application.Documents.Count
For Each oShape In ActiveDocument.InlineShapes
    oShape.Select
        Debug.Print oShape.LinkFormat.SourceFullName
        Debug.Print oShape.LinkFormat.SourcePath
        Debug.Print oShape.LinkFormat.Type
        oShape.Select
        Selection.Copy
        strMsg = strMsg & vbCrLf & oShape.LinkFormat.SourceFullName
 
Next
End Sub

