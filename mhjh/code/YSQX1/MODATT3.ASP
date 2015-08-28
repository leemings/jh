<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
attid=Request.Form("attid")
attackname=Request.Form("attackname")
corp=Request.Form("corp")
condition=Request.Form("condition")
energy=Request.Form("energy")
magic=Request.Form("magic")
attack=Request.Form("attack")
special=Request.Form("special")
introduce=Request.Form("introduce")
attintro=Request.Form("attintro")
opt=Request.Form("opt")
if attid="" or condition="" or energy="" or magic="" or attack="" or introduce="" or attintro="" then Response.Redirect "../error.asp?id=102"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 招式 where id="&attid&" and 门派='"&corp&"'",conn,1,2
if opt="新增"  then rst.AddNew
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="删除" then
rst.Delete
else
rst("招式")=attackname
rst("门派")=corp
rst("修习条件")=condition
rst("消耗精力")=energy
rst("消耗内力")=magic
rst("基本攻击")=attack
rst("特效")=special
rst("说明")=introduce
rst("攻击说明")=attintro
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
msg="<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><div height=100% align=center valign=middle>工作完成<br><a href="&chr(34)&"javascript:location.replace('modatt1.asp?selcorp="&corp&"');"&chr(34)&">返回</a></div></body>"
Response.write msg
%>
