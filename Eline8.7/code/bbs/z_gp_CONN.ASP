<%
	dim gp_conn,gp_connstr,gp_db
	dim sqlbbs,rsbbs
	'更改数据库名字
	gp_db="data/e_stock.asp"		'股票数据库路径，一般不需要修改
	Set gp_conn = Server.CreateObject("ADODB.Connection")
	gp_connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(gp_db)
 	gp_conn.Open gp_connstr
%>