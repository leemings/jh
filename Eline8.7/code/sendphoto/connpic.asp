<%
dim conn,DBPath
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("eline_photopic.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>
