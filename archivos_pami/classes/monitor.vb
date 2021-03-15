Imports System
Imports System.IO
Imports System.Diagnostics

Public Class monitor

    Public watchFolder As FileSystemWatcher
    Dim clsValidate As New validate
    Dim clsMigrate As New migrate
    Dim log As New logTxt

    Public Sub startMonitorA()
        Try
            watchFolder = New System.IO.FileSystemWatcher()
            watchFolder.Path = newFilesPath
            watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.FileName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.Attributes
            AddHandler watchFolder.Created, AddressOf logchange
            watchFolder.EnableRaisingEvents = True
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
        End Try
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        Try
            clsValidate.sArchivoFullPath = e.FullPath
            clsValidate.sArchivo = e.Name
            If clsValidate.validarArchivo() Then
                FileIO.FileSystem.MoveFile(e.FullPath, newFilesPath & "validados\" & e.Name)
            Else
                FileIO.FileSystem.MoveFile(e.FullPath, newFilesPath & "noValidados\" & e.Name)
                log.writeSqlLog(Trim(e.Name), "n", "n")
            End If
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
        End Try
    End Sub

    Public Sub startMonitorB()
        Try
            watchFolder = New System.IO.FileSystemWatcher()
            watchFolder.Path = newFilesPath & "validados\"
            watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.FileName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.Attributes
            AddHandler watchFolder.Created, AddressOf logmigrate
            watchFolder.EnableRaisingEvents = True
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
        End Try
    End Sub

    Private Sub logmigrate(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        Try
            clsMigrate.sArchivoFullPath = e.FullPath
            clsMigrate.importar(e.Name)
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
        End Try
    End Sub

End Class
