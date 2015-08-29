<%
	dim conn1
	dim connstr1
	dim dbbank
	'更改数据库名字
	dbbank="data/e_bank.asp"
	Set conn1 = Server.CreateObject("ADODB.Connection")
	connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbank)
	conn1.Open connstr1
%>