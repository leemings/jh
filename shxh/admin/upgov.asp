<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
usercorp=Session("Ba_jxqy_usercorp")
usergrade=Session("Ba_jxqy_usergrade")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
uname=Trim(Request.Form("uname"))
submit1=Trim(Request.Form("submit1"))
if instr(uname,"'")<>0 then Response.End
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
set rst=Server.CreateObject("adodb.recordset")
rst.Open "select 等级 from 用户 where 姓名='"&uname&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=101"
ugrade=rst("等级")
rst.Close
set rst=nothing
if submit1="升 级" and (usergrade>=Application("Ba_jxqy_exaltgraderight") and ugrade<5 or usergrade=Application("Ba_jxqy_allright")) and usercorp="官府" and usergrade>ugrade Then
	conn.Execute "update 用户 set 等级=等级+1 where 姓名='"&uname&"'"
	msg=uname&"的等级升为"&ugrade+1&"级！"
elseif submit1="降 级" and ((usergrade>=Application("Ba_jxqy_declinegraderight") and ugrade>0) or usergrade=Application("Ba_jxqy_allright")) and usercorp="官府"  and usergrade>ugrade then
	conn.Execute "update 用户 set 等级=等级-1 where 姓名='"&uname&"'"
	msg=uname&"的等级降为"&ugrade-1&"级！"
elseif submit1="开 除" and usercorp="官府" and usergrade=Application("Ba_jxqy_allright") and usergrade>ugrade then
	conn.Execute "update 用户 set 等级=1,门派='无' where 姓名='"&uname&"'"
	msg="解除"&uname&"工作权限完成！"
elseif submit1="聘 用" and usercorp="官府" and usergrade=Application("Ba_jxqy_allright") and usergrade>ugrade then
	conn.Execute "update 用户 set 等级=10,门派='官府' where 姓名='"&uname&"'"
	msg="聘用"&uname&"为官府中人，初始等级为10！"
else 
	msg="无法完成所要求的操作，可能是您的权限不够！"
end if
conn.Close 
set conn=nothing
%>
<head>
<title><%=Application("Ba_jxqy_systemname")%></title>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false >
<p align=center><%=msg%><a href="javascript:location.replace('showgov.asp')" onmouseover="window.status='';return true;" onmouseout="window.status='';return true;">返回</a></p>
</body>