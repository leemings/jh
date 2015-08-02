<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.form("id")
opt=Request.form("submit")
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
rst.Open "select * from biological where id="&id ,conn,1,3
if opt="ÐÂÔö"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="É¾³ý" then
	rst.Delete
else
	rst("biological")=trim(Request.Form("biological"))
	rst("picture")=trim(Request.Form("picture"))
	rst("hp")=clng(trim(Request.Form("hp")))
	rst("attack")=clng(trim(Request.Form("attack")))
	rst("defence")=clng(trim(Request.Form("defence")))
	rst("encourage")=trim(Request.Form("encourage"))
end if	
rst.Update
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showbio.asp"
%>