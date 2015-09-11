<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet  content="text/css">
<title>发表新贴</title>
</head>

<!--#include file="top.asp"-->
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor=#0080C0>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color=#ffffff>爱情江湖　回复</font></td>
  </tr>
<form name="form1" method="post" action="guest_re_save.asp?action=add">
<input name="id" type="hidden" value="<%=request("id")%>">
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center>回复内容</td>
    <td>　<textarea name="re" cols="60" rows="10"></textarea></td>
  </tr>
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center colspan=2><input type="submit" name="Submit" value="回 复">　<input type="reset" name="Submit" value="重 置"></td>
  </tr>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color=#ffffff>发贴者请对自己的言行负责 [<a href=index.asp class=1>返回首页</a>]</font></td>
  </tr>
</form>
</table>
</body>
</html>
