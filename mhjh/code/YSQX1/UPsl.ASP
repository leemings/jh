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
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/mj.asp")
rst.Open "select * from 商店 where id="&id ,conn,1,3
if opt="新增"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="删除" then
rst.Delete
else
rst("物品名")=Trim(Request.Form("物品名"))
rst("类型")=Trim(Request.Form("类型"))
rst("说明")=Trim(Request.Form("说明"))
rst("功能")=Trim(Request.Form("功能"))
rst("精力")=Trim(Request.Form("精力"))
rst("攻击")=Trim(Request.Form("攻击"))
rst("防御")=Trim(Request.Form("防御"))
rst("特效")=Trim(Request.Form("特效"))
rst("会员")=Trim(Request.Form("会员"))
rst("银两")=Clng(Request.Form("银两"))
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showsl.asp"
%>