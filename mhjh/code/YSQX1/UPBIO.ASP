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
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
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
