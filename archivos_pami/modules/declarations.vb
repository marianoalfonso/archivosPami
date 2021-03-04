Module declarations


    Public gsConnectionString As String
    Public newFilesPath As String


    'ESTRUCTURA DONDE SE ALMACENARA EL REGISTRO LEIDO DEL TXT
    Public Structure Registro
        Dim Cuit As String
        Dim RazonSocial As String
        Dim Periodo As String
        Dim Sucursal As String
        Dim Factura As String
        Dim Renglon As Integer
        Dim BeneficioPami As String
        Dim Beneficio As String
        Dim Nombre As String
        Dim Prestacion As String
        Dim Descripcion As String
        Dim Importe As String
    End Structure

End Module
