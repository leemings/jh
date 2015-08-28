<!--#include file="../../config.asp"--><%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 会员,攻击 FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是剑侠中人或者连接超时"
conn.close
response.end
else
hy=rs("会员")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href = "../../exit.asp"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
else
if yx8_mhjh_fellow=false then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href = "javascript:history.back()"
</script>
<%
end if
if rs("攻击")>15000000 then
%>
<script language=vbscript>
MsgBox "错误！你已经是顶级高手了，请回吧！"
location.href = "javascript:history.back()"
</script>
<% else %>
<html>
<head>
<title>第一关</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false >
<table width="562" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
<div align="left"><font color="#FFFFFF"><b>宝藏的第一关是独孤求败的挚友——棋痴张老道</b></font></div>
</td>
<td width="162" rowspan="2" valign="top"><img src="jh1.jpg" width="160" height="250"></td>
</tr>
<tr>
<td width="595" valign="top">
<p style="line-height: 200%; margin: 20"> <font color="#FFFFFF"><font color=red>张老道：</font>好小子，这里是宝藏的第一关，我老人家这把年纪就不跟你动手了，咱们下盘棋，你赢了我就放你过去。但你要输了，我就直接踢你出山口！</font><br><br>
<font color=red><%=username%></font>:<a href="go1.asp">好吧，我一定会下赢你的。</a><br>


</p></td>
</tr>
</table>
</body>
</html>
<% 
end if
end if
end if
end if
rs.Close              
set rs=nothing              
conn.Close              
set conn=nothing 
 %>