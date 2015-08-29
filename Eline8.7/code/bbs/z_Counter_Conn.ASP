<%
dim CntConn,CntConnStr,CntDb
'更改数据库名字
CntDb="data/e_Counter.asp"
Set CntConn = Server.CreateObject("ADODB.Connection")
CntConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(CntDb)
CntConn.Open CntConnStr

function CloseCounterDatabase
	CntConn.close
	Set CntConn = Nothing
End Function
%>
