Module functions

    Function loadConectionString() As Boolean
        Dim configFile As String = Application.StartupPath & "\configuracion_local.txt"
        Try
            'scope = configuracion.SearchValue("AUTHENTICATION", "SCOPE")
            Dim dataSource As String = ConfigGet(configFile, "DATABASE_CREDENTIALS", "DATA_SOURCE")
            Dim initialCatalog As String = ConfigGet(configFile, "DATABASE_CREDENTIALS", "INITIAL_CATALOG")
            Dim userId As String = ConfigGet(configFile, "DATABASE_CREDENTIALS", "USER_ID")
            Dim password As String = ConfigGet(configFile, "DATABASE_CREDENTIALS", "PASSWORD")
            Dim connectionTimeout As String = ConfigGet(configFile, "DATABASE_CREDENTIALS", "CONNECTION_TIMEOUT")
            gsConnectionString = "Data Source=" & dataSource & ";Initial Catalog=" & initialCatalog & ";User ID=" & userId & ";Password=" & password & ";Connection Timeout=" & connectionTimeout
            Return True
        Catch ex As Exception
            MsgBox("error generando la cadena de conexion . . .", vbCritical)
            Return False
        End Try
    End Function

    Function loadGlobalVariables() As Boolean
        Dim configFile As String = Application.StartupPath & "\configuracion_local.txt"
        Try
            newFilesPath = ConfigGet(configFile, "PATHS", "NEW_FILES")
            Return True
        Catch ex As Exception
            MsgBox("error obteniendo variables globales . . .", vbCritical)
            Return False
        End Try
    End Function

    Public Function ConfigGet(ByVal sArchivo As String, ByVal strSection As String, ByVal strItem As String) As String
        'FUNCION: ConfigGet
        'FECHA DE CREACION: 27 de febrero de 2007
        'ULTIMA MODIFICACION: 26 de julio de 2012 (migracion a vb.NET 2008)
        'AUTOR: Mariano Alfonso (genesYs)
        'DESCRIPCION: Dado un encabezado de grupo y un valor, devuelve el
        '             resultado obtenido de un archivo externo de configuracion
        '             parametrizable.
        'PARAMETROS: strSelection --> Cabecera de grupo (String)
        '            strItem      --> Item a buscar (String)
        'DEVOLUCION: (String)
        Try
            'Dim intFile As Integer
            Dim sAux As String
            Dim bytFound As Byte

            Dim sPath As String
            '        sPath = Trim(sArchivo) + ".txt"
            sPath = Trim(sArchivo)
            Dim sFile As New IO.StreamReader(sPath)
            Do While Not sFile.EndOfStream
                sAux = Trim(sFile.ReadLine)
                bytFound = InStr(sAux, strSection)
                If bytFound > 0 Then
                    Exit Do
                End If
            Loop
            If bytFound > 0 Then    'Si encontro la cadena
                bytFound = 0

                Do While Not sFile.EndOfStream
                    sAux = Trim(sFile.ReadLine)
                    bytFound = InStr(sAux, strItem)  'InStr Busca strItem en el archivo
                    If bytFound > 0 Then
                        ConfigGet = Mid(sAux, InStr(sAux, "=") + 1) 'Extrae el resultado
                        Exit Do
                    End If
                Loop
            End If
        Catch ex As Exception
            Err.Clear()
        End Try
    End Function

    Public Function Limpiar_Beneficio(sBeneficio As String) As String
        Dim lBeneficio As Long
        If sBeneficio <> "" Then
            If Len(sBeneficio) = 14 Then
                sBeneficio = Mid(sBeneficio, 4, 8) + Mid(sBeneficio, 13, 2)
                lBeneficio = Val(sBeneficio)
                sBeneficio = Str(lBeneficio)
            Else
                sBeneficio = ""
                Return sBeneficio
            End If
            Return sBeneficio
        Else
            sBeneficio = ""
            Return sBeneficio
        End If
    End Function

    Public Function Limpiar_Cadena(sTexto As String, sCaracter As Char) As String
        Dim auxPos1 As Integer
        Dim auxString1, auxString2 As String
        Try
            If sTexto <> "" Then
                While InStr(sTexto, sCaracter) > 0
                    auxPos1 = InStr(sTexto, sCaracter)
                    auxString1 = Mid(sTexto, 1, auxPos1 - 1)
                    auxString2 = Mid(sTexto, auxPos1 + 1, Len(sTexto))
                    sTexto = auxString1 + auxString2
                End While
                Return sTexto
            Else
                Return sTexto
            End If
        Catch ex As Exception
            sTexto = ""
            Return sTexto
        End Try
    End Function
End Module
