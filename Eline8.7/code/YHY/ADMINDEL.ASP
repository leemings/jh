<%
dim conn,DBPath,sql,sql2,rs_apply


id=request("id")
user_id=request("user_id")

'-----------------------------------------------------------------
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("pic.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
'-----------------------------------------------------------------

set rs_apply=server.createobject("adodb.recordset")
sql="delete * from pic where id =" & id
rs_apply.open sql,conn,3,2


'-----------------------------------------------------------------
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("yhy.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
'-----------------------------------------------------------------
sql2="Update Ãû¼Ë Set photo=photo-1 Where id=" & user_id
rs_apply.open sql2,conn,3,2

  conn.close
  set conn=nothing
  response.redirect "admin.asp?id=" & user_id
   %>