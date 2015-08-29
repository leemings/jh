<%
	dim conncaipiao,connstrcaipiao,caipiaodb
	'更改数据库名字
	caipiaodb="data/e_caipiaodb.asp"                         '彩票数据库路径
	Set conncaipiao = Server.CreateObject("ADODB.Connection")
	connstrcaipiao="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(caipiaodb)
	conncaipiao.Open connstrcaipiao
%>