Option Explicit On
Option Infer On

Module Functions

    Public Function GetTimeStamp(ByVal nPath As String) As String
        Try
            Dim nCheck As New IO.FileInfo(nPath)
            If nCheck.Attributes = IO.FileAttributes.Directory = False Then
                Dim FileInfo As New IO.FileInfo(nPath)
                Return FileInfo.CreationTime
            Else
                Dim DirInfo As New IO.DirectoryInfo(nPath)
                Return DirInfo.CreationTime
            End If
        Catch
            MsgBox("Error: " & Err.Description, MsgBoxStyle.Critical, "TimeStamp Changer")
            Return ""
        End Try
    End Function

    Public Sub SetTimeStamp(ByVal nPath As String, ByVal nDate As String, ByVal nTime As String)
        Try
            'Timestamp format: 01/01/1980 15:45:00
            Dim nCheck As New IO.FileInfo(nPath)
            If nCheck.Attributes = IO.FileAttributes.Directory = False Then
                Dim FileInfo As New IO.FileInfo(nPath)
                FileInfo.CreationTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                FileInfo.LastWriteTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                FileInfo.LastAccessTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
            Else
                Dim DirInfo As New IO.DirectoryInfo(nPath)
                DirInfo.CreationTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                DirInfo.LastWriteTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                DirInfo.LastAccessTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
            End If
            MsgBox("Operation has completed successfully.", MsgBoxStyle.Information, "TimeStamp Changer")
        Catch
            MsgBox("Error: " & Err.Description, MsgBoxStyle.Critical, "TimeStamp Changer")
        End Try
    End Sub

    Public Function FindString(ByVal nText As String) As Boolean
        Dim index As Integer
        For index = 0 To FrmMain.DataGridView1.Rows.Count - 1
            index = index
            If FrmMain.DataGridView1.Rows.Item(index).Cells(0).Value = nText Then
                Return True
            End If
        Next
        Return False
    End Function

End Module
