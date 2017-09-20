Imports System.ComponentModel

Public Class frmMain

    Private xFileHandler As New FileHandler
    Private xCleaner As EmailCleaner
    Private xSelectedFilePath As String = String.Empty

    Public Delegate Sub DelegateUpdateForm(ByVal text As String)
    Public Delegate Sub DelegateSetProgressBarMaximum(ByVal Max As Integer)
    Public Delegate Sub DelegateIncrementProgressBar()
    Public Delegate Sub DelegateCloseForm()
    Public Delegate Sub DelegateUpdateFile(ByVal filename As String)

    Private xWorker_UpdateFormText As New DelegateUpdateForm(AddressOf UpdateFormText)
    Private xWorker_SetProgressMax As New DelegateSetProgressBarMaximum(AddressOf SetProgressBarMaximum)
    Private xWorker_IncrementProgess As New DelegateIncrementProgressBar(AddressOf IncrementProgressBar)
    Private xWorker_CloseForm As New DelegateCloseForm(AddressOf CloseForm)
    Private xWorker_UpdateFile As New DelegateUpdateFile(AddressOf UpdateFile)

    Public Sub RunTest()
        Dim testList As New List(Of EmailTargetInfo)
        testList.Add(New EmailTargetInfo With {.Subfolder = "Sent Items", .Subject = "code review for ticket#14618"})

        xCleaner.PerformCleanup(testList)
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        xCleaner = New EmailCleaner
        xCleaner.IncrementProgressCall = AddressOf IncrementProgressBar
        xCleaner.SetStatusCall = AddressOf UpdateFormText
        xCleaner.StartProgressCall = AddressOf SetProgressBarMaximum

        lblFilename.Text = "Select folder where input Excel files are located to get started"
        btnStart.Enabled = False
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        bgWorker.RunWorkerAsync()
    End Sub

    Private Sub bgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles bgWorker.DoWork
        Main()
        'RunTest()
    End Sub

    Public Sub Main()
        Dim fileCollection As ObjectModel.ReadOnlyCollection(Of String) =
            My.Computer.FileSystem.GetFiles(xSelectedFilePath, FileIO.SearchOption.SearchTopLevelOnly, "*.xl*")

        For Each filename As String In fileCollection
            xWorker_UpdateFile(filename)

            'parse file
            xWorker_UpdateFormText("parsing file data")
            Dim targetList As List(Of EmailTargetInfo) = ParseFile(filename)

            'process parsed targets
            Dim cleanupResults As List(Of CleanupResult)
            cleanupResults = xCleaner.PerformCleanup(targetList)

            'write results log
            xWorker_UpdateFormText("writing results to log")
            WriteResultsLog(filename, cleanupResults)
        Next
    End Sub

    Private Function ParseFile(ByVal fileName As String) As List(Of EmailTargetInfo)
        Dim parsedCollection As New List(Of EmailTargetInfo)

        Try
            parsedCollection = xFileHandler.ParseFile(fileName)

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        End Try

        Return parsedCollection
    End Function

    Private Sub WriteResultsLog(ByVal filename As String, ByVal results As List(Of CleanupResult))
        Try
            xFileHandler.WriteCleanupResults(filename, results)

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        End Try
    End Sub

    Private Sub btnSetFolder_Click(sender As Object, e As EventArgs) Handles btnSetFolder.Click
        FolderBrowser.ShowDialog()
        xSelectedFilePath = FolderBrowser.SelectedPath
        lblFilename.Text = "Selected folder: " & xSelectedFilePath
        lblStatus.Text = "Click start to begin processing"
        btnStart.Enabled = True
    End Sub

    Private Sub SetProgressBarMaximum(ByVal Max As Integer)
        If InvokeRequired Then
            Dim worker As New DelegateSetProgressBarMaximum(AddressOf SetProgressBarMaximum)
            Invoke(worker, New Object() {Max})
        Else
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = Max
            Refresh()
            My.Application.DoEvents()
        End If
    End Sub

    Private Sub IncrementProgressBar()
        If InvokeRequired Then
            Dim worker As New DelegateIncrementProgressBar(AddressOf IncrementProgressBar)
            Invoke(worker, New Object() {})
        Else
            ProgressBar1.Increment(1)
            Refresh()
            My.Application.DoEvents()
        End If
    End Sub

    Private Sub UpdateFormText(ByVal text As String)
        If InvokeRequired Then
            Dim worker As New DelegateUpdateForm(AddressOf UpdateFormText)
            Invoke(worker, New Object() {text})
        Else
            lblStatus.Text = "Status: " & text
            Refresh()
            My.Application.DoEvents()
        End If
    End Sub

    Private Sub UpdateFile(ByVal filename As String)
        If InvokeRequired Then
            Dim worker As New DelegateUpdateFile(AddressOf UpdateFile)
            Invoke(worker, New Object() {filename})
        Else
            lblFilename.Text = "File: " & filename
            Refresh()
            My.Application.DoEvents()
        End If
    End Sub

    Private Sub CloseForm()
        If InvokeRequired Then
            Dim worker As New DelegateCloseForm(AddressOf CloseForm)
            Invoke(worker, New Object() {})
        Else
            Close()
        End If
    End Sub
End Class
