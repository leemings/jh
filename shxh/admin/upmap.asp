<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.form("id")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="¹Ù¸®") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
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