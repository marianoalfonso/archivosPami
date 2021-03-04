Imports System.Data.SqlClient
Imports System
Imports System.IO
Public Class migrate

    Public sArchivoFullPath As String

    Public Sub importar(fileName As String)
        Dim sRazonSocial As String
        Dim sCuit As String
        Dim sPeriodo As String
        Dim sSucursal As String
        Dim sFactura As String
        Dim sRenglon As Integer
        Dim sImporte As String
        Dim sRegistro(0) As Registro
        Dim sLinea As String
        Dim iContador, iAux As Integer
        Dim rValor1, rValor2 As Double
        Dim iLimiteSuperiorArray As Integer

        Dim registroGrabado As Boolean = False
        Dim log As New logTxt
        log.logPath = fileName

        Try
            iContador = 0
            'Dim sPath As String = "c:\ArchivosFacturacion\" + sArchivo
            Dim sPath As String = sArchivoFullPath
            Dim sContent As String = vbNullString

            'OBTENEMOS LOS DATOS DE CABECERA
            sRazonSocial = ConfigGet(sArchivoFullPath, "HEADER", "PRESTADOR")
            sCuit = ConfigGet(sArchivoFullPath, "HEADER", "CUIT")
            sCuit = Limpiar_Cadena(sCuit, "-")
            sPeriodo = ConfigGet(sArchivoFullPath, "HEADER", "PERIODOFACTURADO")
            sPeriodo = Limpiar_Cadena(sPeriodo, "-")
            sSucursal = ConfigGet(sArchivoFullPath, "HEADER", "NUMEROFACTURA")
            sFactura = ConfigGet(sArchivoFullPath, "HEADER", "NUMEROFACTURA")

            Using sFile As New IO.StreamReader(sPath)
                sLinea = Trim(sFile.ReadLine)
                While sLinea <> "[BOTTOM]"
                    sRenglon = 1

                    If sLinea = "[DETAIL]" Or sLinea = "endline" Then   'si es "DETAIL[ o "endline" asume que a continuación hay datos
                        sRegistro(iContador).Beneficio = Trim(sFile.ReadLine)   'leo la siguiente linea
                        If sRegistro(iContador).Beneficio = "[BOTTOM]" Then 'si la linea leida es "BOTTOM" es el final del archivo y salimos del proceso
                            Exit While
                        Else    'si la linea leida NO es "BOTTOM" hay datos de consumo
                            sRegistro(iContador).BeneficioPami = sRegistro(iContador).Beneficio
                            sRegistro(iContador).Beneficio = Limpiar_Beneficio(sRegistro(iContador).Beneficio)
                        End If
                        sRegistro(iContador).RazonSocial = sRazonSocial
                        sRegistro(iContador).Cuit = sCuit
                        sRegistro(iContador).Periodo = sPeriodo
                        sRegistro(iContador).Sucursal = sSucursal
                        sRegistro(iContador).Factura = sFactura
                        sRegistro(iContador).Renglon = sRenglon
                        sRegistro(iContador).Nombre = Trim(sFile.ReadLine)
                        sRegistro(iContador).Prestacion = Trim(sFile.ReadLine)
                        sRegistro(iContador).Descripcion = Trim(sFile.ReadLine)
                        sImporte = Trim(sFile.ReadLine)
                        iLimiteSuperiorArray = UBound(sRegistro)
                        'CHEQUEAMOS EXISTENCIA
                        If iLimiteSuperiorArray > 0 Then
                            iAux = 0
                            Do Until iAux = iLimiteSuperiorArray
                                If sRegistro(iAux).Beneficio = sRegistro(iContador).Beneficio Then
                                    If sRegistro(iAux).Prestacion = sRegistro(iContador).Prestacion Then
                                        rValor1 = Val(Replace(sRegistro(iContador).Importe, ",", "."))
                                        rValor2 = Val(Replace(sRegistro(iAux).Importe, ",", "."))
                                        sRegistro(iAux).Importe = (rValor1 + rValor2).ToString
                                        ReDim Preserve sRegistro(UBound(sRegistro) - 1)
                                        iContador -= 1
                                        Exit Do
                                    Else
                                        sRegistro(iContador).Importe = sImporte
                                    End If
                                Else
                                    sRegistro(iContador).Importe = sImporte
                                End If
                                iAux += 1
                            Loop
                        Else
                            sRegistro(iContador).Importe = sImporte
                        End If
                        iContador += 1
                        ReDim Preserve sRegistro(UBound(sRegistro) + 1)
                    End If
                    sLinea = Trim(sFile.ReadLine)
                    'Application.DoEvents()
                End While
                'we clear the last position of the array that is empty
                ReDim Preserve sRegistro(UBound(sRegistro) - 1)
            End Using
            If GrabarFacturacion(sRegistro, "D") Then
                log.writeLog("archivo importado a la base de datos")
                log.writeSqlLog(Trim(fileName), "s", "s")
                'System.IO.File.Move(newFilesPath & "validados\" & fileName, newFilesPath & "validados\migrados\" & fileName)
            Else
                log.writeLog("error importando archivo a la base de datos")
                'System.IO.File.Move(newFilesPath & "\validados\" & fileName, newFilesPath & "noMigrados\" & fileName)
                log.writeSqlLog(Trim(fileName), "s", "n")
            End If
            log = Nothing
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            log = Nothing
        End Try
    End Sub

    Private Function GrabarFacturacion(ByVal sArray() As Registro, sArea As Char) As Boolean
        Dim Elementos, iElemento As Integer
        Dim Conn As New SqlConnection(gsConnectionString)
        Try
            Elementos = sArray.Length - 1
            For iElemento = 0 To Elementos
                Dim Cmd As New SqlCommand
                Cmd.CommandText = "SP_IMPORTAR_FACTURACION"
                Cmd.CommandType = CommandType.StoredProcedure
                Cmd.Connection = Conn
                Cmd.Parameters.AddWithValue("@AREA", Trim(sArea))
                Cmd.Parameters.AddWithValue("@CUIT", Trim(sArray(iElemento).Cuit))
                Cmd.Parameters.AddWithValue("@RAZON_SOCIAL", Trim(sArray(iElemento).RazonSocial))
                Cmd.Parameters.AddWithValue("@PERIODO", Trim(sArray(iElemento).Periodo))
                Cmd.Parameters.AddWithValue("@SUCURSAL", Trim(sArray(iElemento).Sucursal))
                Cmd.Parameters.AddWithValue("@FACTURA", Trim(sArray(iElemento).Factura))
                Cmd.Parameters.AddWithValue("@RENGLON", Trim(sArray(iElemento).Renglon))
                Cmd.Parameters.AddWithValue("@BENEFICIOPAMI", Trim(sArray(iElemento).BeneficioPami))
                Cmd.Parameters.AddWithValue("@BENEFICIO", Trim(sArray(iElemento).Beneficio))
                Cmd.Parameters.AddWithValue("@NOMBRE", Trim(sArray(iElemento).Nombre))
                If Trim(sArray(iElemento).Prestacion) = "Geriatria" Or Trim(sArray(iElemento).Prestacion) = "Hemodialisis" Then
                    sArray(iElemento).Prestacion = "Prestacion"
                End If
                Cmd.Parameters.AddWithValue("@PRESTACION", Trim(sArray(iElemento).Prestacion))
                Cmd.Parameters.AddWithValue("@DESCRIPCION", Trim(sArray(iElemento).Descripcion))
                Cmd.Parameters.AddWithValue("@IMPORTE", Trim(Replace(sArray(iElemento).Importe, ",", ".")))
                Conn.Open()
                Cmd.ExecuteScalar()
                Conn.Close()
            Next iElemento
            Conn = Nothing
            Return True
        Catch ex As Exception
            If Err.Number = 5 Then
                MsgBox("El archivo seleccionado ya fue importado", vbCritical)
                Err.Clear()
                Return False
            End If
        End Try
    End Function

End Class
