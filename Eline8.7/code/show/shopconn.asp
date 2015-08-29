<%
	dim connshop
	dim strshop
	dim dbshop
	'更改社区超市据库名字
	dbshop="../bbs/data/eline_bbs_6.3.0.asp"
	Set connshop = Server.CreateObject("ADODB.Connection")
	strshop="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbshop)
	'如果你的服务器采用较老版本Access驱动，请用下面连接方法
	'strshop="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(dbshop)	
	connshop.Open strshop%>