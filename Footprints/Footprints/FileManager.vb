Imports System.IO

Public Class FileManager
    Function readFile(ByVal filePath As String) As List(Of String)
        Dim records As New List(Of String)

        Try
            Dim sr As StreamReader = New StreamReader(filePath)

            Do While sr.Peek() >= 0
                Dim recordLine = sr.ReadLine()
                records.Add(recordLine)
            Loop
        Catch
            MsgBox("Unable to read the file " & filePath, MsgBoxStyle.OkOnly, "Footprints Viewer")
        End Try

        Return records
    End Function

End Class
