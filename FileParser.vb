Imports OfficeOpenXml
Imports System.IO
Imports Core

Public Class FileHandler

    Public Function ParseFile(ByVal fileName As String) As List(Of EmailTargetInfo)
        Const INPUT_WORKSHEET_NAME As String = "Outlook File Details"
        Dim recordList As New List(Of EmailTargetInfo)
        Dim package As ExcelPackage

        Try
            Dim fInfo As New FileInfo(fileName)
            package = New ExcelPackage(fInfo)

            If package Is Nothing Then
                Throw New Exception("Could not load excel package")
            End If

            Dim workbook As ExcelWorkbook = package.Workbook
            Dim sheets As ExcelWorksheets = workbook.Worksheets
            Dim sheet As ExcelWorksheet = sheets(INPUT_WORKSHEET_NAME)

            If sheet Is Nothing Then
                Throw New Exception("Could not read worksheet. Attempted to find sheet name " & INPUT_WORKSHEET_NAME)
            End If

            Dim moreRows As Boolean = True
            Dim rowNumber As Integer = 0
            Dim firstColumnValueCurrentRow As String
            Dim firstColumnValueNextRow As String
            Dim outlookFilename As String = String.Empty

            Dim targetItem As EmailTargetInfo

            While moreRows
                rowNumber += 1

                firstColumnValueCurrentRow = If(sheet?.Cells(rowNumber, 1).Text, String.Empty)
                firstColumnValueNextRow = If(sheet?.Cells(rowNumber + 1, 1).Text, String.Empty)

                If String.IsNullOrWhiteSpace(firstColumnValueCurrentRow) Then
                    If String.IsNullOrWhiteSpace(firstColumnValueNextRow) Then
                        moreRows = False 'there can be one blank row between each mailbox, 2 blanks should mean EOF
                    End If

                    Continue While
                End If

                If firstColumnValueCurrentRow.Contains(".pst") Then
                    outlookFilename = firstColumnValueCurrentRow
                    rowNumber += 1 'jump past header for next mailbox data
                    Continue While
                End If

                targetItem = New EmailTargetInfo
                targetItem.PstFilename = outlookFilename
                targetItem.Subfolder = If(sheet?.Cells(rowNumber, 2).Text, String.Empty)
                targetItem.Subject = If(sheet?.Cells(rowNumber, 4).Text, String.Empty)

                recordList.Add(targetItem)
            End While

        Catch ex As Exception
            'Logger.Instance.LogError(ex, "filename: " & fileName)
        Finally
            package?.Dispose()
        End Try

        Return recordList
    End Function

    Public Sub WriteCleanupResults(ByVal fileName As String, ByVal resultsCollection As List(Of CleanupResult))
        Const RESULTS_WORKSHEET_NAME As String = "Email Cleanup Results"
        Dim package As ExcelPackage

        Try
            Dim fInfo As New FileInfo(fileName)
            package = New ExcelPackage(fInfo)

            If package Is Nothing Then
                Throw New Exception("Could not load excel package")
            End If

            Dim workbook As ExcelWorkbook = package.Workbook
            Dim sheets As ExcelWorksheets = workbook.Worksheets

            sheets.Add(RESULTS_WORKSHEET_NAME)

            Dim sheet As ExcelWorksheet = sheets(RESULTS_WORKSHEET_NAME)

            If sheet Is Nothing Then
                Throw New Exception("Could not create and find worksheet. Attempted to find sheet name " & RESULTS_WORKSHEET_NAME)
            End If

            Dim rowNumber As Integer = 0

            sheet.Cells(rowNumber, 1).Value = "Mailbox"
            sheet.Cells(rowNumber, 2).Value = "Subfolder"
            sheet.Cells(rowNumber, 3).Value = "Subject"
            sheet.Cells(rowNumber, 4).Value = "Completed"
            sheet.Cells(rowNumber, 5).Value = "Error Text"

            For Each result As CleanupResult In resultsCollection
                rowNumber += 1
                sheet.Cells(rowNumber, 1).Value = result.Target.PstFilename
                sheet.Cells(rowNumber, 2).Value = result.Target.Subfolder
                sheet.Cells(rowNumber, 3).Value = result.Target.Subject
                sheet.Cells(rowNumber, 4).Value = result.Completed.ToString
                sheet.Cells(rowNumber, 5).Value = result.ErrorText
            Next

            package.Save()

        Catch ex As Exception
            'Logger.Instance.LogError(ex, "filename: " & fileName)
        Finally
            package?.Dispose()
        End Try

    End Sub

End Class
