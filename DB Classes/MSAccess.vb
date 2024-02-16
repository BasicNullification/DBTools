Imports System.Data.OleDb
Imports System.Runtime.InteropServices.WindowsRuntime

<ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsDual)>
<Guid("62D63BD4-27B8-4121-A9A5-287FA0158582")>
Public Interface IMSAccess
    Inherits IBaseDB
    Function NewRecord(ByVal tableName As String) As NewRecord
End Interface

<ProgId("DBTools.MSAccess"), ComVisible(True), ClassInterface(ClassInterfaceType.None)>
<Guid("63B9EFEE-3818-4263-89A6-80FEF2882960")>
Public Class MSAccess

    Inherits BaseDB
    Implements IMSAccess

    Public Function NewRecord(ByVal tableName As String) As NewRecord Implements IMSAccess.NewRecord
        Return New NewRecord(tableName, ConnectionString)
    End Function

End Class