Public Class Form1
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        If loadConectionString() Then
            If loadGlobalVariables() Then
                startMonitorA()
                startMonitorB()
            End If
        Else
            MessageBox.Show("error obteniendo datos de configuracion")
        End If
        Application.Exit()
    End Sub

    'scan root folder for previous files
    Sub scanFolder()
        Dim clsValidate As New validate
        Dim log As New logTxt
        Try
            Dim fileInfo As System.IO.FileInfo
            For Each foundfile As String In My.Computer.FileSystem.GetFiles(newFilesPath, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
                fileInfo = My.Computer.FileSystem.GetFileInfo(foundfile)
                clsValidate.sArchivoFullPath = foundfile
                clsValidate.sArchivo = fileInfo.Name
                If clsValidate.validarArchivo() Then
                    FileIO.FileSystem.MoveFile(foundfile, newFilesPath & "validados\" & fileInfo.Name)
                Else
                    FileIO.FileSystem.MoveFile(foundfile, newFilesPath & "noValidados\" & fileInfo.Name)
                    log.writeSqlLog(Trim(fileInfo.Name), "n", "n")
                End If
            Next
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
        Finally
            clsValidate = Nothing
            log = Nothing
        End Try
    End Sub

    'watch the folder 'archivosfacturacion' to validate them
    Private Sub startMonitorA()
        Try
            Dim clsMonitor As New monitor
            clsMonitor.startMonitorA()
            Me.Label1.ForeColor = Color.Green
            Me.Label1.Text = "file monitor A started"
        Catch ex As Exception
            Me.Label1.ForeColor = Color.Red
            Me.Label1.Text = "error starting the file monitor A"
        End Try
    End Sub

    'watch the folder 'validated files' to migrate them
    Private Sub startMonitorB()
        Try
            Dim clsMonitor As New monitor
            clsMonitor.startMonitorB()
            Me.Label2.ForeColor = Color.Green
            Me.Label2.Text = "file monitor B started"
        Catch ex As Exception
            Me.Label2.ForeColor = Color.Red
            Me.Label2.Text = "error starting the file monitor B"
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            scanFolder()
        Catch ex As Exception
            MessageBox.Show("la aplicacion no pudo iniciarse")
            Application.Exit()
        End Try
    End Sub
End Class
