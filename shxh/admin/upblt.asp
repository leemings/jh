<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.form("id")
pid=Request.Form("pid")
opt=Request.form("submit")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
ballottxt=Request.Form("ballottxt")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballot where id="&id&" and pid="&pid ,conn,1,2
if opt="����"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="ɾ��" then
	rst.Delete
else
	rst("pid")=clng(pid)
	rst("parti")=trim(Request.Form("parti"))
	rst("vote")=clng(trim(Request.Form("vote")))
end if	
rst.Update
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showblt.asp?pid="&pid
%>