<%
	dim conn
	dim dbpath
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("../51e_djmusic.asp")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>