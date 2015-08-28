

<%
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=016"
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("mlimlihong.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
id=Request.QueryString("id")
delsql="delete from data where id="&id
conn.execute(delsql)
conn.close
response.redirect"add.asp"
%>
