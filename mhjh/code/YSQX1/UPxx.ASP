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
rst.Open "select * from mj where id="&id ,conn,1,3
if opt="新增"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="删除" then
rst.Delete
else
rst("名称")=Trim(Request.Form("名称"))
rst("类型")=Trim(Request.Form("类型"))
rst("所有者")=Trim(Request.Form("所有者"))
rst("攻击")=Clng(Request.Form("攻击"))
rst("防御")=Clng(Request.Form("防御"))
rst("速度")=Clng(Request.Form("速度"))
rst("特效")=Trim(Request.Form("特效"))
rst("等级")=Clng(Request.Form("等级"))
rst("图")=Trim(Request.Form("图"))
rst("时间")=Trim(Request.Form("时间"))
rst("状态")=Trim(Request.Form("状态"))
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Redirect "showxx.asp"
%>