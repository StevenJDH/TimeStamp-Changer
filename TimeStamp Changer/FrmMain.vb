Option Explicit On
Option Infer On

Public Class FrmMain
    Inherits System.Windows.Forms.Form

    Private Sub FileAddToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileAddToolStripMenuItem1.Click
        Dim Results As DialogResult
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "All files (*.*)|*.*"
        Results = OpenFileDialog1.ShowDialog
        If Results = DialogResult.OK Then
            If CDate(GetTimeStamp(OpenFileDialog1.FileName)).Year < CDate("1/1/1980").Year Then
                MsgBox("Error: '" & OpenFileDialog1.SafeFileName & "' has a date older than 1/1/1980 which is not supported and cannot be changed.", MsgBoxStyle.Critical, "TimeStamp Changer")
                Exit Sub
            End If
            If FindString(OpenFileDialog1.SafeFileName) = False Then
                Dim n As Integer = DataGridView1.Rows.Add
                DataGridView1.Rows.Item(n).Cells(0).Value = OpenFileDialog1.SafeFileName
                DataGridView1.Rows.Item(n).Cells(1).Value = OpenFileDialog1.FileName
                RemoveSelectedItemToolStripMenuItem.Enabled = True
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
                If DataGridView1.Rows.Count > 1 And BatchToolStripMenuItem1.Checked = False And BatchToolStripMenuItem1.Tag = "" Then
                    BatchToolStripMenuItem1.Tag = "1"
                    Results = MsgBox("You are adding more than one item to the list. Would you like to change to Batch Mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "TimeStamp Changer")
                    If Results = DialogResult.Yes Then
                        BatchToolStripMenuItem1.Checked = True
                        Button1_Click(Nothing, Nothing)
                    End If
                End If
                Column1.HeaderText = "Drag & Drop Files/Folders [Total: " & DataGridView1.Rows.Count & "]"
            End If
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If BackgroundWorker1.IsBusy = True And Button2.Text = "Cancel" Then
            BackgroundWorker1.CancelAsync()
            Exit Sub
        End If
        Dim Results As DialogResult
        Results = MsgBox("Are you sure you want to change the timestamp(s)?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "TimeStamp Changer")
        If Results = DialogResult.No Then
            Exit Sub
        End If
        Select Case BatchToolStripMenuItem1.Checked
            Case "False"
                Button2.Text = "Changing..."
                Button2.Enabled = False
                Button1.Enabled = False
                RemoveSelectedItemToolStripMenuItem.Enabled = False
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = False
                DatePicker1.Enabled = False
                TimePicker1.Enabled = False
                FileToolStripMenuItem.Enabled = False
                SetTimeStamp(DataGridView1.Rows.Item(DataGridView1.CurrentCellAddress.Y).Cells(1).Value, DatePicker1.Value.Date, TimePicker1.Value)
                FileToolStripMenuItem.Enabled = True
                TimePicker1.Enabled = True
                DatePicker1.Enabled = True
                RemoveSelectedItemToolStripMenuItem.Enabled = True
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
                Button2.Text = "Change"
            Case "True"
                Button2.Text = "Cancel"
                Button1.Enabled = False
                RemoveSelectedItemToolStripMenuItem.Enabled = False
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = False
                DatePicker1.Enabled = False
                TimePicker1.Enabled = False
                FileToolStripMenuItem.Enabled = False
                ProgressBar1.Maximum = DataGridView1.Rows.Count
                BackgroundWorker1.RunWorkerAsync()
        End Select
    End Sub

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DatePicker1.MaxDate = Now
        TimePicker1.MaxDate = Now
    End Sub

    Private Sub DatePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DatePicker1.ValueChanged
        If DatePicker1.Value.Date = Today.Date Then
            TimePicker1.MaxDate = Now
        Else
            TimePicker1.MaxDate = "12/31/9998" 'This allow the numbers to go higher than curret time.
        End If
    End Sub

    Private Sub FolderAddToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FolderAddToolStripMenuItem.Click
        Dim Results As DialogResult
        FolderBrowserDialog1.Description = "Select a folder to change its' timestamp:"
        FolderBrowserDialog1.SelectedPath = ""
        Results = FolderBrowserDialog1.ShowDialog()
        If Results = DialogResult.OK Then
            If Microsoft.VisualBasic.Right(FolderBrowserDialog1.SelectedPath, 2) = ":\" Then
                MsgBox("Only folders and not drives maybe selected. Please try again.", MsgBoxStyle.Information, "TimeStamp Changer")
                Exit Sub
            End If
            Dim CurrentDir As New IO.FileInfo(FolderBrowserDialog1.SelectedPath & "\")
            If CDate(GetTimeStamp(FolderBrowserDialog1.SelectedPath)).Year < CDate("1/1/1980").Year Then
                MsgBox("Error: '" & CurrentDir.Directory.Name & "' has a date older than 1/1/1980 which is not supported and cannot be changed.", MsgBoxStyle.Critical, "TimeStamp Changer")
                Exit Sub
            End If
            If FindString(CurrentDir.Directory.Name) = False Then
                Dim n As Integer = DataGridView1.Rows.Add
                DataGridView1.Rows.Item(n).Cells(0).Value = CurrentDir.Directory.Name
                DataGridView1.Rows.Item(n).Cells(1).Value = FolderBrowserDialog1.SelectedPath & "\"
                RemoveSelectedItemToolStripMenuItem.Enabled = True
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
                If DataGridView1.Rows.Count > 1 And BatchToolStripMenuItem1.Checked = False And BatchToolStripMenuItem1.Tag = "" Then
                    BatchToolStripMenuItem1.Tag = "1"
                    Results = MsgBox("You are adding more than one item to the list. Would you like to change to Batch Mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "TimeStamp Changer")
                    If Results = DialogResult.Yes Then
                        BatchToolStripMenuItem1.Checked = True
                        Button1_Click(Nothing, Nothing)
                    End If
                End If
                Column1.HeaderText = "Drag & Drop Files/Folders [Total: " & DataGridView1.Rows.Count & "]"
            End If
        End If
    End Sub
    
    Public Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim nResetInfo As String
        nResetInfo = Now
        DatePicker1.MaxDate = nResetInfo
        DatePicker1.Value = nResetInfo
        TimePicker1.MaxDate = nResetInfo
        TimePicker1.Value = nResetInfo
    End Sub

    Private Sub BatchToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BatchToolStripMenuItem1.Click
        If BatchToolStripMenuItem1.Checked = True Then
            BatchToolStripMenuItem1.Checked = False
        Else
            BatchToolStripMenuItem1.Checked = True
        End If
        Button1_Click(Nothing, Nothing)
    End Sub

    Private Sub Donate5PaypalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Donate5PaypalToolStripMenuItem.Click
        On Error Resume Next
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=8493677")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("TimeStamp Changer 1.0 (25/2/2010)" & vbNewLine & vbNewLine & "Author: Steven Jenkins De Haro" & _
        vbNewLine & "A Steve Creation/Convergence" & vbNewLine & vbNewLine & _
        "Microsoft .NET Framework 3.5", MsgBoxStyle.OkOnly, "TimeStamp Changer")
    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        RemoveSelectedItemToolStripMenuItem.Enabled = False
        RemoveSelectedItemMenuToolStripMenuItem.Enabled = False
        BatchToolStripMenuItem1.Tag = ""
        DataGridView1.Rows.Clear()
        Column1.HeaderText = "Drag & Drop Files/Folders"
        Button1_Click(Nothing, Nothing)
        Button1.Enabled = False
        Button2.Enabled = False
        DataGridView1.Focus()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count < 1 Then
            Exit Sub
        End If
        If BatchToolStripMenuItem1.Checked = False Then 'If not in Batch mode show individual file info
            Try
                Dim nCreationTime As String = GetTimeStamp(DataGridView1.Rows.Item(DataGridView1.CurrentCellAddress.Y).Cells(1).Value)
                If Not nCreationTime = "" Then
                    DatePicker1.Value = nCreationTime
                    TimePicker1.Value = nCreationTime
                Else
                    Button1_Click(Nothing, Nothing)
                End If
            Catch
                MsgBox("Error: Could not determine the item's current timestamp. The current date and time will be used.", MsgBoxStyle.Critical, "TimeStamp Changer")
                Button1_Click(Nothing, Nothing)
            End Try
        End If
    End Sub

    Private Sub DataGridView1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragDrop
        Dim Results As DialogResult
        Try
            Dim nDropedFiles As String() = DirectCast(e.Data.GetData(DataFormats.FileDrop), String())
            For Each nFiles As String In nDropedFiles
                If Microsoft.VisualBasic.Right(nFiles, 2) = ":\" Then
                    MsgBox("Only folders and not drives maybe selected. Please try again.", MsgBoxStyle.Information, "TimeStamp Changer")
                    Exit Sub
                End If
                Dim nCheck As New IO.FileInfo(nFiles)
                If nCheck.Attributes = IO.FileAttributes.Directory = False Then
                    If CDate(GetTimeStamp(nFiles)).Year < CDate("1/1/1980").Year Then
                        MsgBox("Error: '" & IO.Path.GetFileName(nFiles) & "' has a date older than 1/1/1980 which is not supported and cannot be changed.", MsgBoxStyle.Critical, "TimeStamp Changer")
                    Else
                        If FindString(IO.Path.GetFileName(nFiles)) = False Then
                            Dim n As Integer = DataGridView1.Rows.Add
                            DataGridView1.Rows.Item(n).Cells(0).Value = IO.Path.GetFileName(nFiles)
                            DataGridView1.Rows.Item(n).Cells(1).Value = nFiles
                            Column1.HeaderText = "Drag & Drop Files/Folders [Total: " & DataGridView1.Rows.Count & "]"
                            RemoveSelectedItemToolStripMenuItem.Enabled = True
                            RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
                            Button1.Enabled = True
                            Button2.Enabled = True
                            If DataGridView1.Rows.Count > 1 And BatchToolStripMenuItem1.Checked = False And BatchToolStripMenuItem1.Tag = "" Then
                                BatchToolStripMenuItem1.Tag = "1"
                                Results = MsgBox("You are adding more than one item to the list. Would you like to change to Batch Mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "TimeStamp Changer")
                                If Results = DialogResult.Yes Then
                                    BatchToolStripMenuItem1.Checked = True
                                    Button1_Click(Nothing, Nothing)
                                End If
                            End If
                        End If
                    End If
                Else
                    Dim CurrentDir As New IO.FileInfo(nFiles & "\")
                    If CDate(GetTimeStamp(nFiles & "\")).Year < CDate("1/1/1980").Year Then
                        MsgBox("Error: '" & CurrentDir.Directory.Name & "' has a date older than 1/1/1980 which is not supported and cannot be changed.", MsgBoxStyle.Critical, "TimeStamp Changer")
                    Else
                        If FindString(CurrentDir.Directory.Name) = False Then
                            Dim n As Integer = DataGridView1.Rows.Add
                            DataGridView1.Rows.Item(n).Cells(0).Value = CurrentDir.Directory.Name
                            DataGridView1.Rows.Item(n).Cells(1).Value = nFiles & "\"
                            Column1.HeaderText = "Drag & Drop Files/Folders [Total: " & DataGridView1.Rows.Count & "]"
                            RemoveSelectedItemToolStripMenuItem.Enabled = True
                            RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
                            Button1.Enabled = True
                            Button2.Enabled = True
                            If DataGridView1.Rows.Count > 1 And BatchToolStripMenuItem1.Checked = False And BatchToolStripMenuItem1.Tag = "" Then
                                BatchToolStripMenuItem1.Tag = "1"
                                Results = MsgBox("You are adding more than one item to the list. Would you like to change to Batch Mode?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "TimeStamp Changer")
                                If Results = DialogResult.Yes Then
                                    BatchToolStripMenuItem1.Checked = True
                                    Button1_Click(Nothing, Nothing)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Catch
            MsgBox("Error: " & Err.Description, MsgBoxStyle.Critical, "TimeStamp Changer")
        End Try
    End Sub

    Private Sub DataGridView1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DataGridView1.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim nDate As String = DatePicker1.Value.Date
        Dim nTime As String = TimePicker1.Value
        Dim index As Integer
        For index = 0 To DataGridView1.Rows.Count - 1
            index = index
            If BackgroundWorker1.CancellationPending = True Then
                e.Result = "Cancel"
                Exit Sub
            End If
            DataGridView1.Rows(index).Selected = True
            Try
                'Timestamp format: 01/01/1980 15:45:00
                Dim nCheck As New IO.FileInfo(DataGridView1.Rows.Item(index).Cells(1).Value)
                If nCheck.Attributes = IO.FileAttributes.Directory = False Then
                    Dim FileInfo As New IO.FileInfo(DataGridView1.Rows.Item(index).Cells(1).Value)
                    FileInfo.CreationTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                    FileInfo.LastWriteTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                    FileInfo.LastAccessTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                Else
                    Dim DirInfo As New IO.DirectoryInfo(DataGridView1.Rows.Item(index).Cells(1).Value)
                    DirInfo.CreationTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                    DirInfo.LastWriteTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                    DirInfo.LastAccessTime = DateTime.Parse(nDate & " " & Format(CDate(nTime), "HH:mm:ss"))
                End If
                BackgroundWorker1.ReportProgress(index + 1)
            Catch
                e.Result = Err.Description
            End Try
        Next
        Threading.Thread.Sleep(2000) 'This is to allow progress bar to catch up when finished.
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Result = "Cancel" Then
            MsgBox("Operation was canceled. Only " & ProgressBar1.Value & " out of " & DataGridView1.Rows.Count & " items were changed.", MsgBoxStyle.Information, "TimeStamp Changer")
        ElseIf e.Result IsNot Nothing Then
            MsgBox("Operation was completed unsuccessfully. The following item had problems and others might have been affected because of it:" & vbNewLine & vbNewLine & DataGridView1.Rows.Item(ProgressBar1.Value).Cells(0).Value & _
                vbNewLine & vbNewLine & "Check if above item is not set to Ready-only and that you have sufficient permissions to access it. Then rerun TimeStampe Changer to try again and maybe fix those files that now may say 'how long ago' for a timestamp and not a time.", MsgBoxStyle.Critical, "TimeStamp Changer")
        Else
            MsgBox("Operation has completed successfully.", MsgBoxStyle.Information, "TimeStamp Changer")
        End If
        ProgressBar1.Value = 0
        FileToolStripMenuItem.Enabled = True
        TimePicker1.Enabled = True
        DatePicker1.Enabled = True
        Button1.Enabled = True
        RemoveSelectedItemToolStripMenuItem.Enabled = True
        RemoveSelectedItemMenuToolStripMenuItem.Enabled = True
        Button2.Text = "Change"
    End Sub

    Private Sub RemoveSelectedItemToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RemoveSelectedItemToolStripMenuItem.Click
        If DataGridView1.Rows.Count > 0 And Button2.Text = "Change" Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            If DataGridView1.Rows.Count < 1 Then
                RemoveSelectedItemToolStripMenuItem.Enabled = False
                RemoveSelectedItemMenuToolStripMenuItem.Enabled = False
                Column1.HeaderText = "Drag & Drop Files/Folders"
                Button1_Click(Nothing, Nothing)
                Button1.Enabled = False
                Button2.Enabled = False
                DataGridView1.Focus()
                Exit Sub
            End If
            Column1.HeaderText = "Drag & Drop Files/Folders [Total: " & DataGridView1.Rows.Count & "]"
        End If
    End Sub

    Private Sub RemoveSelectedItemMenuToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RemoveSelectedItemMenuToolStripMenuItem.Click
        RemoveSelectedItemToolStripMenuItem_Click(Nothing, Nothing)
    End Sub
End Class

