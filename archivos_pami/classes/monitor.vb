Imports System
Imports System.IO
Imports System.Diagnostics

Public Class monitor

    Public watchFolder As FileSystemWatcher
    Dim clsValidate As New validate
    Dim clsMigrate As New migrate
    Dim log As New logTxt

    Public Sub startMonitorA()
        watchFolder = New System.IO.FileSystemWatcher()
        'watchFolder.Path = "C:\ArchivosFacturacionX\"
        watchFolder.Path = newFilesPath
        watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.FileName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.Attributes
        AddHandler watchFolder.Created, AddressOf logchange
        watchFolder.EnableRaisingEvents = True
    End Sub

    Private Sub logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        clsValidate.sArchivoFullPath = e.FullPath
        clsValidate.sArchivo = e.Name
        If clsValidate.validarArchivo() Then
            FileIO.FileSystem.MoveFile(e.FullPath, newFilesPath & "validados\" & e.Name)
            'System.IO.File.Move(e.FullPath, newFilesPath & "validados\" & e.Name)
            'System.IO.File.Move(e.FullPath, "C:\ArchivosFacturacionX\validados\" & e.Name)
        Else
            FileIO.FileSystem.MoveFile(e.FullPath, newFilesPath & "noValidados\" & e.Name)
            Log.writeSqlLog(Trim(e.Name), "n", "n")
            'System.IO.File.Move(e.FullPath, newFilesPath & "noValidados\" & e.Name)
            'System.IO.File.Move(e.FullPath, "C:\ArchivosFacturacionX\noValidados\" & e.Name)
        End If
    End Sub

    Public Sub startMonitorB()
        watchFolder = New System.IO.FileSystemWatcher()
        watchFolder.Path = newFilesPath & "validados\"
        watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.FileName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or IO.NotifyFilters.Attributes
        AddHandler watchFolder.Created, AddressOf logmigrate
        watchFolder.EnableRaisingEvents = True
    End Sub

    Private Sub logmigrate(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        clsMigrate.sArchivoFullPath = e.FullPath
        clsMigrate.importar(e.Name)
    End Sub

End Class
