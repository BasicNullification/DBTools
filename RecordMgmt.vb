Imports System.Data.OleDb

Public MustInherit Class RecordMgmt
    Implements IDisposable

    Private ReadOnly _connStr As String
    Protected ReadOnly _tblName As String
    Protected _kvPairs As New Dictionary(Of String, Object)

    Protected ReadOnly Property ConnObj As OleDbConnection

    Protected Sub New(tableName As String, connStr As String)
        _tblName = tableName
        _connStr = connStr
        ConnObj = New OleDbConnection(_connStr)
    End Sub

    Protected Function Execute(cmdText As String, Optional parameters As List(Of OleDbParameter) = Nothing) As Integer
        Dim retVal As Integer
        Using ConnObj
            ConnObj.Open()
            Using c As New OleDb.OleDbCommand(cmdText, ConnObj)
                If parameters IsNot Nothing Then
                    For Each param As OleDbParameter In parameters
                        c.Parameters.Add(param)
                    Next
                End If
                retVal = c.ExecuteNonQuery()
            End Using
        End Using
        Return retVal
    End Function

    Public Overridable Sub Dispose() Implements IDisposable.Dispose
        ConnObj.Dispose()
    End Sub
End Class

Public Interface INewRecord
    WriteOnly Property Field(ByVal FieldName As String) As Object
End Interface

Public Class NewRecord
    Inherits RecordMgmt
    Implements INewRecord

    Friend Sub New(tableName As String, connStr As String)
        MyBase.New(tableName, connStr)
    End Sub

    Public WriteOnly Property Field(ByVal FieldName As String) As Object Implements INewRecord.Field
        Set(value As Object)
            If _kvPairs.ContainsKey(FieldName) Then
                _kvPairs(FieldName) = value     'Update the value if the key exists
            Else
                _kvPairs.Add(FieldName, value)  '.. or add a new key-value pair
            End If
        End Set
    End Property

    Public Function Write() As Integer
        Dim fields = String.Join(", ", _kvPairs.Keys)
        Dim values = String.Join(", ", _kvPairs.Keys.Select(Function(k) "?"))
        Dim cmdText = $"INSERT INTO {_tblName} ({fields}) VALUES ({values})"

        Dim parameters = _kvPairs.Values.Select(Function(v) New OleDbParameter With {.Value = v}).ToList()

        Return MyBase.Execute(cmdText, parameters)
    End Function

End Class