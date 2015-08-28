<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.form("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"¹Ù¸®" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from map where id="&id ,conn,1,3
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
rst("displace")=Trim(Request.Form("displace"))
rst("affair")=Trim(Request.Form("affair"))
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showmap.asp"
%>
