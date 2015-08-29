<%
	dim conn
	dim dbpath
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("showimg.asp")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>