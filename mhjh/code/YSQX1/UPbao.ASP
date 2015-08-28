<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
id=Request.form("id")
opt=Request.form("submit")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
ballottxt=Request.Form("ballottxt")
set conn=server.CreateObject("adodb.connection")
set rst=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/jihe.asp")
rst.Open "select * from 宝物 where id="&id ,conn,1,3
if opt="新增"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="删除" then
rst.Delete
else
rst("防御力")=Trim(Request.Form("防御力"))
rst("价格")=Trim(Request.Form("价格"))
rst("攻击力")=Trim(Request.Form("攻击力"))
rst("拿走")=Clng(Request.Form("拿走"))
rst("宝物名")=Trim(Request.Form("宝物名"))
rst("时间")=Trim(Request.Form("时间"))
rst("获宝者")=Trim(Request.Form("获宝者"))
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showbao.asp"
%>