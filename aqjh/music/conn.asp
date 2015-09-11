<%
startime=timer()
dim conn
dim connstr
set conn=server.CreateObject("adodb.connection")
DBPath = Server.MapPath("music_mdb.asp")
conn.open "provider=microsoft.jet.oledb.4.0; data source="&DBpath

function CloseDB

Conn.Close
set Conn=Nothing

End Function
%>