<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
corpid=Request.Form("corpid")
corp=Request.Form("corp")
earnings=Request.Form("earnings")
introduce=Request.Form("introduce")
condition=Request.Form("condition")
chaton=cbool(Request.Form("chaton"))
opt=Request.Form("opt")
if corpid="" or corp="" or earnings="" or introduce="" or condition="" or opt="" then Response.Redirect "..\errpr.asp?id=102"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 门派  where id="&corpid,conn,1,2
if opt="新增"  then
rst.AddNew
end if
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
if opt="删除" then
rst.Delete
else
rst("门派")=corp
rst("工资系数")=earnings
rst("简介")=introduce
rst("加入条件")=condition
rst("chaton")=chaton
end if
rst.Update
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.write "<head><link rel='stylesheet' href='../chatroom/css.css'>.</head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><div height=100% align=center valign=middle>工作完成<br><a href="&chr(34)&"javascript:location.replace('modcorp1.asp');"&chr(34)&">返回</a></div></body>"
%>
