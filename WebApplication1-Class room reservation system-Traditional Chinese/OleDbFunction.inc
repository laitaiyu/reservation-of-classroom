<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<Script Languate="VB" Runat="Server">
  'strSQL 為 SQL 敘述，FileName 為資料庫檔名，TableName 為資料表名稱。
  Function CreateDataSet(strSQL As String, FileName As String, TableName As String) As DataSet
    Dim ConnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                               "Data Source=" & Server.MapPath(FileName)
    Dim objCmd As New OleDbDataAdapter(strSQL, ConnString)
    Dim DS As New DataSet()
    objCmd.Fill(DS, TableName)
    CreateDataSet = DS
  End Function
</Script>