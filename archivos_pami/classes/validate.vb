Public Class validate

    Public sArchivoFullPath As String
    Public sArchivo As String
    Public sError As String

    Public Function validarArchivo() As Boolean
        Dim log As New logTxt
        log.logPath = sArchivo
        Dim sLinea As String
        Dim renglones As Int16 = 0
        Dim sPath As String = sArchivoFullPath
        Try
            log.writeLog(sArchivo)
            log.writeLog(Now)
            Using sFile As New IO.StreamReader(sPath)
                sLinea = Trim(sFile.ReadLine)
                If sLinea <> "[STATUS]" Then
                    log.writeLog("no se encuentra la linea [STATUS]")
                    Return False
                End If
                sLinea = Trim(sFile.ReadLine)
                If sLinea <> "SF=closed" Then
                    log.writeLog("no se encuentra la linea SF=closed")
                    Return False
                End If
                sLinea = Trim(sFile.ReadLine)
                If sLinea <> "[HEADER]" Then
                    log.writeLog("no se encuentra la linea [HEADER]")
                    Return False
                End If
                sLinea = Left(Trim(sFile.ReadLine), 10)
                If sLinea <> "PRESTADOR=" Then
                    log.writeLog("no se encuentra la linea PRESTADOR=")
                    Return False
                Else
                    If ConfigGet(sArchivoFullPath, "HEADER", "PRESTADOR") = "" Then
                        log.writeLog("la linea PRESTADOR tiene un valor nulo")
                        sFile.Close()
                        sFile.Dispose()
                        Return False
                    End If
                End If
                sLinea = Left(Trim(sFile.ReadLine), 5)
                If sLinea <> "CUIT=" Then
                    log.writeLog("no se encuentra la linea CUIT=")
                    Return False
                Else
                    If ConfigGet(sArchivoFullPath, "HEADER", "CUIT") = "" Then
                        log.writeLog("la linea CUIT tiene un valor nulo")
                        Return False
                    End If
                End If
                sLinea = Left(Trim(sFile.ReadLine), 17)
                If sLinea <> "PERIODOFACTURADO=" Then
                    log.writeLog("no se encuentra la linea PERIODOFACTURADO=")
                    Return False
                Else
                    If ConfigGet(sArchivoFullPath, "HEADER", "PERIODOFACTURADO") = "" Then
                        log.writeLog("la linea PERIODOFACTURADO tiene un valor nulo")
                        Return False
                    End If
                End If
                sLinea = Left(Trim(sFile.ReadLine), 14)
                If sLinea <> "NUMEROFACTURA=" Then
                    log.writeLog("no se encuentra la linea NUMEROFACTURA=")
                    Return False
                Else
                    If ConfigGet(sArchivoFullPath, "HEADER", "NUMEROFACTURA") = "" Then
                        log.writeLog("la linea NUMEROFACTURA tiene un valor nulo")
                        Return False
                    End If
                End If
                sLinea = Trim(sFile.ReadLine)
                If sLinea <> "[DETAIL]" Then
                    log.writeLog("no se encuentra la linea [DETAIL]")
                    Return False
                End If
                sLinea = Trim(sFile.ReadLine)
                Do While sLinea <> "[BOTTOM]"
                    If sLinea <> "endline" And sLinea <> "[BOTTOM]" Then
                        renglones = renglones + 1
                        If renglones > 5 Then
                            log.writeLog("error en algun bloque de detalle de consumo o falta la linea [BOTTOM]")
                            Return False
                        End If
                    Else
                        If renglones < 5 Then
                            log.writeLog("error en algun bloque de detalle de consumo o falta la linea [BOTTOM]")
                            Return False
                        End If
                        renglones = 0
                    End If
                    sLinea = Trim(sFile.ReadLine)
                Loop
                log.writeLog("archivo validado correctamente")
                Return True
            End Using
        Catch ex As Exception
            log.writeSystemErrorLog(Err.Number & "-" & Err.Description)
            Return False
        End Try
    End Function

End Class
