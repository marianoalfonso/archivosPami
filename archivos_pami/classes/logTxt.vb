Imports System.IO
Imports System.Data.SqlClient

Public Class logTxt

    Public logPath As String
    Public archivo As String
    Public validado As String
    Public importado As String

    Public Sub writeLog(message As String)
        Dim writer As StreamWriter
        Try
            writer = File.AppendText(newFilesPath & "log\(" & Now.ToString("yyyyMMdd") & ") - " & logPath)
            'writer = File.AppendText("C:\ArchivosFacturacionX\log\(" & Now.ToString("yyyyMMdd") & ") - " & logPath)
            writer.Write(message & vbCrLf)
            writer.Flush()
            writer.Close()
            writer.Dispose()
        Catch ex As Exception
            'MessageBox.Show("error escribiendo en el log")
        End Try
    End Sub

    Public Sub writeSystemErrorLog(errorMessage As String)
        Dim writer As StreamWriter
        Try
            writer = File.AppendText(newFilesPath & "systemErrorLog\systemErrorLog.txt")
            writer.Write(Now.ToString("yyyyMMdd") & ") - " & errorMessage & vbCrLf)
            writer.Flush()
            writer.Close()
            writer.Dispose()
        Catch ex As Exception
            'MessageBox.Show("error escribiendo en el log")
        End Try
    End Sub

    Public Sub writeSqlLog(archivo As String, validado As String, importado As String)
        Dim conn As New SqlConnection(gsConnectionString)
        Try
            Dim cmd As New SqlCommand
            cmd.CommandText = "mpa_01000"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = conn
            cmd.Parameters.AddWithValue("@archivo", Trim(archivo))
            cmd.Parameters.AddWithValue("@fecha", Now.ToString("yyyyMMdd"))
            cmd.Parameters.AddWithValue("@validado", Trim(validado))
            cmd.Parameters.AddWithValue("@importado", Trim(importado))
            conn.Open()
            cmd.ExecuteScalar()
            conn.Close()
            conn = Nothing
        Catch ex As Exception
            conn = Nothing
        End Try
    End Sub


End Class
