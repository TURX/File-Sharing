Attribute VB_Name = "Save"
Public Function SaveFileFromRes(vntresourceId As Variant, sTYPE As String, sfilename As String) As Boolean
Dim bytimage() As Byte
Dim ifilenum As Integer
On Error GoTo SaveFileFromRes_err
    SaveFileFromRes = True
    bytimage = LoadResData(vntresourceId, sTYPE)
    ifilenum = FreeFile
Open sfilename For Binary As ifilenum
Put #ifilenum, , bytimage
Close ifilenum
SaveFileFromRes_err: Exit Function
SaveFileFromRes = False: Exit Function
End Function
