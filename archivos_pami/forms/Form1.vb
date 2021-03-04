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
End Class
