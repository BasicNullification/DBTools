<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("877531E0-69A0-40DA-8512-99AFDE73B088")>
Public Interface IBaseDB
    Property Provider As String
    Property DBFilePath As String
End Interface

<ProgId("DBTools.BaseDB"), ComVisible(True)>
<Guid("B9CD05BF-60FB-425D-A1F9-EDDAD4E65490")>
Public MustInherit Class BaseDB

    Implements IBaseDB

    Private _ConnectionStr As String = vbNullString
    Private _ConnectionStrOverride As Boolean = False
    Private _Provider As String = "Microsoft.ACE.OleDb.12.0"
    Private _dbPath As String = vbNullString
    'Private ReadOnly _ConnectionObj As OleDb.OleDbConnection

    Public Property Provider As String Implements IBaseDB.Provider
        Get
            Return _Provider
        End Get
        Set(value As String)
            _Provider = value
        End Set
    End Property

    Public Property DBFilePath As String Implements IBaseDB.DBFilePath
        Get
            Return _dbPath
        End Get
        Set(value As String)
            _dbPath = value
        End Set
    End Property

    Public Property ConnectionString As String
        Get
            If _ConnectionStrOverride Then
                Return _ConnectionStr
            Else
                Return $"Provider={Provider};Data Source={_dbPath};"
            End If
        End Get
        Set(value As String)
            _ConnectionStrOverride = True
            _ConnectionStr = value
        End Set
    End Property

End Class
