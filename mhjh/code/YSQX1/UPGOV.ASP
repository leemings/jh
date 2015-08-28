<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
uname=Trim(Request.Form("uname"))
submit1=Trim(Request.Form("submit1"))
if instr(uname,"'")<>0 then Response.End
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("yx8_mhjh_connstr")
set rst=Server.CreateObject("adodb.recordset")
rst.Open "select 等级 from 用户 where 姓名='"&uname&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
ugrade=rst("等级")
rst.Close
set rst=nothing
if submit1="升 级" Then
conn.Execute "update 用户 set 等级=等级+1 where 姓名='"&uname&"'"
msg=uname&"的等级升为"&ugrade+1&"级！"
elseif submit1="降 级" and ugrade>0 then
conn.Execute "update 用户 set 等级=等级-1 where 姓名='"&uname&"'"
msg=uname&"的等级降为"&ugrade-1&"级！"
elseif submit1="开 除" then
conn.Execute "update 用户 set 等级=1,门派='无' where 姓名='"&uname&"'"
msg="解除"&uname&"工作权限完成！"
elseif submit1="聘 用" then
conn.Execute "update 用户 set 等级=40,门派='官府' where 姓名='"&uname&"'"
msg="聘用"&uname&"为官府中人，初始等级为40！"
else
msg="无法完成所要求的操作，可能是您的权限不够！"
end if
conn.Close
set conn=nothing
%>
<head>
</head>
<body  background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false >
<p align=center><%=msg%><a href="javascript:location.replace('showgov.asp')" onmouseover="window.status='';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>
