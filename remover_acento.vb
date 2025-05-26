Function Acento(Caract As String)

Dim A As String
Dim B As String
Dim i As Integer
Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

For i = 1 To Len(AccChars)
    A = Mid(AccChars, i, 1)
    B = Mid(RegChars, i, 1)
    Caract = Replace(Caract, A, B)
    Caract = Replace(Caract, "'", "")
    Caract = Replace(Caract, ".", "")
    Caract = Replace(Caract, "-", "")
Next
Acento = Caract
End Function