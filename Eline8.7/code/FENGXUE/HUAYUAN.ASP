
<!--#include file="func213.asp"-->
<!--#include file="sub.asp" -->
<%
Response.Buffer=true
if session("sjjh_name")="" then Response.Redirect "../../error.asp?id=210"
if session("advtime")<>"" then
if session("advtime")>now()-0.0001 then  Response.Redirect "../../error.asp?id=490"
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
if rs("体力")<0 or rs("状态")="死亡" then Response.Redirect "../../exit.asp"
tl=rs("体力")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#ffffff;
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
p{font-size:16;color:#ffffff;}
</style>
<link rel="stylesheet" href="../../style4.css" type="text/css">
</head>
<body  bgproperties="fixed" bgcolor="#000000" vlink="#000000" text="#000000" oncontextmenu=self.event.returnValue=false background="christmas1.jpg">
<center>
<%
if tl<20 then
%>
<script language=vbscript>
MsgBox "你疲劳程度已超出范围，为了您的身体，还是离开民居为上！"
location.href = "javascript:history.back()"
</script>
<%
else
message=huayuan(my,sh,sy)

end if
end if%>
  <table width="100" border="0" cellspacing="1" cellpadding="4" bgcolor="#990000">
    <tr>
<td background="../../main/zuo.gif">　</td>
<td>
        <table border=0 align=center width=350 cellpadding="0" cellspacing="0">
          <tr>
<td>
              <div align="center">
                <center>
              <table height="260" width="100%" bgcolor="#000000" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
                <tr>
<td height="37">
                    <div align="center"><font color="#ffffff"><strong><font color="#FF0000" size="-1">发生事件</font></strong></font></div>

<tr>
<td height="182" valign="top"><%=message%><br>
<br>
</td>
</tr>
<td align=center height="29">
<div align="right">
<p align="center">
<input type=button value="返回" onClick="javascript:history.go(-1)" name="button" class="input">
</div>
</td>
</table>
                </center>
              </div>
</td>
</tr>
</table>
</td>
<td background="../../main/you.gif">　</td>
</tr>
<tr>
<td></td>
<td background="../../main/xiabg.gif">　</td>
<td></td>
</tr>
</table>
<font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1">Copyright   
  &copy; 2003-2004 All rights   
  reserved</font> 
</center> 
</body> 
</html>