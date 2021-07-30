Option Explicit
Function enterQuery(fName As String, sqlStr As String)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim connStr As String
        Dim providerStr As String
        'Dim sqlStr As String
        'Dim connStr As String
        Dim fields() As Variant
        Dim arrLen As Integer
        Dim vLen As Integer
        Dim cLen As Integer
        Dim n As Integer
        Dim see As String
        Dim i As Integer
        
        'fd.InitialFileName = ThisWorkbook.Path
        connStr = "Data Source=" & fName
        providerStr = "Microsoft.ACE.OLEDB.12.0"
        With cn
            .ConnectionString = connStr
            .Provider = providerStr
            .Open
        End With
        
        'sqlStr = "SELECT 'First Name','Last Name', 'Email Address', City FROM Students"
        'sqlStr = "SELECT [First Name], [Last Name], [E-mail Address], City FROM Students"

        rs.Open sqlStr, cn
        
        see = rs.GetString
        rs.MoveFirst
        fields = rs.GetRows
        'arrLen = UBound(fields(0)) - LBound(fields(0)) + 1
        'arrLen = UBound(Application.Transpose(Application.Index(fields, 0, 0))) - LBound(Application.Transpose(Application.Index(fields, 0, 0))) + 1
        cLen = UBound(fields, 1) - LBound(fields, 1) + 1
        vLen = UBound(fields, 2) - LBound(fields, 2) + 1
        n = 0
        enterQuery = fields
    End Function
