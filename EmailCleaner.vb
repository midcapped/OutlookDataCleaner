Imports Microsoft.Office.Interop
Imports Core
Imports System.Runtime.InteropServices

Public Class EmailCleaner

    Private xRunMode As RunMode = RunMode.TestingOnly_NoActions
    Private xApp As Outlook.Application
    Private xExplorer As Outlook.Explorer
    Private xNamespace As Outlook.NameSpace

    Const CLEANUP_FOLDER_NAME As String = "Laptop Cleanup Items" 'note: folder needs to be manually created in current state
    Private xCleanupFolder As Outlook.MAPIFolder

    Public Delegate Sub StartProgress(ByVal total As Integer)
    Public Delegate Sub IncrementProgress()
    Public Delegate Sub SetStatus(ByVal status As String)

    Private xStartProgress As StartProgress
    Private xIncrementProgress As IncrementProgress
    Private xSetStatus As SetStatus

    Public WriteOnly Property StartProgressCall As StartProgress
        Set(value As StartProgress)
            xStartProgress = value
        End Set
    End Property

    Public WriteOnly Property IncrementProgressCall As IncrementProgress
        Set(value As IncrementProgress)
            xIncrementProgress = value
        End Set
    End Property

    Public WriteOnly Property SetStatusCall As SetStatus
        Set(value As SetStatus)
            xSetStatus = value
        End Set
    End Property

    Public Sub New()
        xApp = New Outlook.Application
        xNamespace = xApp.GetNamespace("MAPI")
        xExplorer = xApp.ActiveExplorer
    End Sub

    Public Function PerformCleanup(ByVal targets As List(Of EmailTargetInfo)) As List(Of CleanupResult)

        Dim results As New List(Of CleanupResult) With {.Capacity = targets.Count}

        Try
            xStartProgress(targets.Count)

            For Each target As EmailTargetInfo In targets

                xSetStatus("Finding item " & target.Subject)
                Dim matchedItem As Outlook.MailItem = FindMailItem(target)

                If matchedItem Is Nothing Then
                    results.Add(New CleanupResult With {.Completed = False, .ErrorText = "Failed to find item"})
                    Continue For
                End If

                xSetStatus("Processing matched item " & target.Subject)

                If Not ProcessMatchedItem(matchedItem) Then
                    results.Add(New CleanupResult With {.Completed = False, .ErrorText = "Failed to process matched item"})
                    Continue For
                End If

                results.Add(New CleanupResult With {.Completed = True})

                xIncrementProgress()
            Next

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        End Try

        xSetStatus("cleanup complete")

        Return results
    End Function

    Private Function FindMailItem(ByVal target As EmailTargetInfo) As Outlook.MailItem
        Dim targetItem As Outlook.MailItem = Nothing
        Dim items As Outlook.Items

        Try
            If Not SetFolder(target) Then
                Throw New Exception("Unable to set target folder. Folder name: " & target.Subfolder)
            End If

            items = xExplorer.CurrentFolder.Items

            For i As Integer = 1 To items.Count - 1
                If TypeOf (items(i)) Is Outlook.MailItem Then
                    Dim mailItem As Outlook.MailItem = CType(items(i), Outlook.MailItem)

                    If String.Equals(target.Subject, mailItem.Subject, StringComparison.OrdinalIgnoreCase) Then
                        targetItem = mailItem
                        Exit For
                    End If

                    ReleaseComObject(mailItem)
                End If
            Next

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        Finally
            ReleaseComObject(items)
        End Try

        Return targetItem
    End Function

    Private Function ProcessMatchedItem(ByVal item As Outlook.MailItem) As Boolean
        Try

            If xRunMode = RunMode.DeleteMatchedItems Then
                item.Delete()

                'ElseIf xRunMode = RunMode.MoveMatchesToCleanupFolder Then
                '    If xCleanupFolder Is Nothing Then
                '        SetFolder(CLEANUP_FOLDER_NAME)
                '        xCleanupFolder = xExplorer.CurrentFolder
                '    End If

                '    If xCleanupFolder IsNot Nothing Then
                '        item.Move(xCleanupFolder)
                '    Else
                '        Throw New Exception("Could locate folder to move matched item, item not processed.")
                '    End If
            End If

        Catch ex As Exception
            'Logger.Instance.LogError(ex, "item subject: " & item.Subject)
        End Try

        Return True
    End Function

    ''' <summary>
    ''' Note: this method isn't working, could be permissions error. 
    ''' </summary>
    ''' <param name="folderName"></param>
    ''' <returns></returns>
    Private Function CreateFolder(ByVal folderName As String) As Boolean
        Dim stores As Outlook.Stores
        Dim store As Outlook.Store
        Dim rootFolder As Outlook.MAPIFolder

        Try
            If xCleanupFolder IsNot Nothing Then
                Exit Try
            End If

            stores = xNamespace.Stores
            store = stores(1) 'TODO: switch stores based on mailbox name? or should new folder always be in first mailbox?
            rootFolder = store.GetRootFolder

            xCleanupFolder = rootFolder.Folders.Add(folderName, Outlook.OlDefaultFolders.olFolderInbox)

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        Finally
            ReleaseComObject(store)
            ReleaseComObject(stores)
            ReleaseComObject(rootFolder)
        End Try

        Return xCleanupFolder IsNot Nothing
    End Function

    Private Function SetFolder(ByVal target As EmailTargetInfo) As Boolean
        Dim stores As Outlook.Stores
        Dim store As Outlook.Store
        Dim rootFolder As Outlook.MAPIFolder
        Dim folders As Outlook.Folders
        Dim folder As Outlook.MAPIFolder


        Try
            Dim matchedFile As Boolean

            stores = xNamespace.Stores
            For i As Integer = 1 To stores.Count - 1
                store = stores(i)

                If String.Equals(store.FilePath, target.PstFilename) Then
                    matchedFile = True
                    Exit For
                End If
            Next

            If Not matchedFile Then
                Throw New Exception("Could not locate Outlook store with matching pst file name. Searched for: " & target.PstFilename)
            End If

            rootFolder = store.GetRootFolder
            folders = rootFolder.Folders

            For Each f As Outlook.OlDefaultFolders In [Enum].GetValues(GetType(Outlook.OlDefaultFolders))
                folder = xNamespace.GetDefaultFolder(f)

                If String.Equals(target.Subfolder, folder.Name, StringComparison.OrdinalIgnoreCase) Then
                    xExplorer.CurrentFolder = folder
                    Return True
                End If
            Next

            For i As Integer = 1 To folders.Count - 1
                folder = folders(i)

                If String.Equals(target.Subfolder, folder.Name, StringComparison.OrdinalIgnoreCase) Then
                    xExplorer.CurrentFolder = folder
                    Return True
                End If
            Next

        Catch ex As Exception
            'Logger.Instance.LogError(ex)
        Finally
            ReleaseComObject(store)
            ReleaseComObject(stores)
            ReleaseComObject(rootFolder)
            ReleaseComObject(folders)
            ReleaseComObject(folder)
        End Try

        Return False
    End Function

    Private Sub ReleaseComObject(ByVal o As Object)
        If o IsNot Nothing Then
            Marshal.ReleaseComObject(o)
        End If
    End Sub

    Private Enum RunMode As Integer
        DeleteMatchedItems
        MoveMatchesToCleanupFolder
        TestingOnly_NoActions
    End Enum

End Class
