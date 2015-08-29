<%
bgcolor=Application("sjjh_backgroundcolor")
bgimage=Application("sjjh_backgroundimage")
%>
<!--#include file="func213.asp"-->
<!--#include file="sub.asp" -->
<%
Response.Buffer=true
if session("sjjh_name")="" then Response.Redirect "../error.asp?id=016"
if session("advtime")<>"" then
if session("advtime")>now()-0.0006 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
sh=request.form("sh")
sy=request.form("sy")
if request.form("h")="1" then
my=session("sjjh_name")
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="select * from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs("体力")<0 or rs("状态")="死" then 
session.Abandon
end if
tl=rs("体力")
%>
<html>
<head>
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
<LINK href="../style.css" rel=stylesheet></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<center>
<%
if tl<20 then
%>
<script language=vbscript>
MsgBox "你疲劳程度已超出范围，为防不测，还是离开孤岛为上！"
location.href = "javascript:history.back()"
</script>
<%
else
message=huayuan(my,sh,sy)

end if
end if%>
<table border=1 align=center width=350 cellpadding="0" cellspacing="1" bordercolor="#000000" bgcolor="#FFFFFF">
<tr>
<td>
<table height="225" align="center" width="100%">
<tr>
<td height="37">
<div align="center"><strong><font color="#FF0000">突发事件</font></strong></div>

<tr>
<td height="147" valign="top">
  <p align="center">  <img border="0" src="../chat/img/251.GIF"><br><%=message%>

</p>
</td>
</tr>
<td align=center height="29">
<div align="right">
<p align="center">
<input type=button value="返 回 孤 岛" onClick="location.href='index.asp'" name="button">&nbsp; 
<font color="#FF0000"><b>[ </b></font> <a href="javascript:self.close()"><font color="#FF0000"><b>离开孤岛</b></font></a><font color="#FF0000"><b> ]</b></font> 
</div> 
</td> 
</table> 
</td> 
</tr> 
</table> 
</center> 
</body> 
</html>