<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.form("id")
opt=Request.form("submit")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"¹Ù¸®" then Response.Redirect "../exit.asp"
ballottxt=Request.Form("ballottxt")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballotsystem where id="&id ,conn,1,2
if opt="ÐÂÔö"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="É¾³ý" then
conn.Execute "delete from ballot where pid="&id
rst.Delete
else
rst("title")=Trim(Request.Form("title"))
rst("explain")=Trim(Request.Form("explain"))
rst("deadline")=cdate(Request.Form("deadline"))
rst("condition")=Trim(Request.Form("condition"))
rst("active")=cbool(Request.Form("active"))
rst("voter")=server.HTMLEncode(Trim(Request.Form("voter")))
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showballot.asp"
%>
