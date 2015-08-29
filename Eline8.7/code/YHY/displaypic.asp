<% 
id=request("id")
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("Eline_yhy_pic.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
set rs=server.createobject("ADODB.recordset") 
sql="select * from pic where id=" & id
rs.open sql,conn,1,1 
Response.ContentType = "image/jpeg" 
Response.BinaryWrite rs("big")
rs.close 
set rs=nothing 
set connGraph=nothing 

%> 
