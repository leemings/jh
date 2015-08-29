<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("Eline_yhy_pic.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>