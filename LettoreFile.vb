Imports System.IO
Imports System.String
Public Class LettoreFile
    Private FileInfo
    ''''funzione per lettura directory (numero file e nome dei file)
    Public Function ReadFileFromFolder(ByVal strPercorso As String, ByVal i As Integer) As String

        Dim objDirInfo As DirectoryInfo

        objDirInfo = New DirectoryInfo(strPercorso)

        FileInfo = objDirInfo.GetFiles

        If i >= 0 Then
            Dim FileName As String = FileInfo(i).name()
            ReadFileFromFolder = FileName
        Else
            Dim nFile As Integer = objDirInfo.GetFiles.Length
            ReadFileFromFolder = nFile
        End If

    End Function
End Class
