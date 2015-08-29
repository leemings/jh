<%
set conn=server.CreateObject("adodb.connection")
path=server.MapPath("wish22.mdb")
conn.open "provider=microsoft.jet.oledb.4.0; data source="&path&""
%>