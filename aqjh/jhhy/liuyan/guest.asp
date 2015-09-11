<%
user=session("aqjh_name")
if user="" then
response.write"<script>alert('对不起，只有江湖玩家才可以在这里发布信息给站长，请您去登陆江湖之后再来！'); history.back();</script>"
response.end()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet  content="text/css">
<title>报告事情</title>
</head>
<!--#include file="top.asp"-->
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor=#0080C0>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color="#ffffff">有什么事可以在这里写给站长看哦!</font></td>
  </tr>
<form name="form1" method="post" action="guest_save.asp?action=add">
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center>报告主题</td>
    <td >　<input name="title" type="text" maxlength="48" size=100></td>
  </tr>
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center>报告内容</td>
    <td >　<textarea name="book" cols="85" rows="10"></textarea></td>
  </tr>
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center colspan=2><input type="submit" name="Submit" value="发 布">　<input type="reset" name="Submit" value="重 写"></td>
  </tr>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color=#ffffff>发贴者请对自己的言行负责 
      </font></td>
  </tr>
</form>
</table>
</html>
