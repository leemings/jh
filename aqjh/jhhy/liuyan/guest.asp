<%
user=session("aqjh_name")
if user="" then
response.write"<script>alert('�Բ���ֻ�н�����Ҳſ��������﷢����Ϣ��վ��������ȥ��½����֮��������'); history.back();</script>"
response.end()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet  content="text/css">
<title>��������</title>
</head>
<!--#include file="top.asp"-->
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor=#0080C0>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color="#ffffff">��ʲô�¿���������д��վ����Ŷ!</font></td>
  </tr>
<form name="form1" method="post" action="guest_save.asp?action=add">
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center>��������</td>
    <td >��<input name="title" type="text" maxlength="48" size=100></td>
  </tr>
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center>��������</td>
    <td >��<textarea name="book" cols="85" rows="10"></textarea></td>
  </tr>
  <tr bgcolor="#F7F9FF">
    <td height=30 align=center colspan=2><input type="submit" name="Submit" value="�� ��">��<input type="reset" name="Submit" value="�� д"></td>
  </tr>
  <tr>
    <td bgcolor="#0080C0" height=20 align=center colspan=2><font color=#ffffff>����������Լ������и��� 
      </font></td>
  </tr>
</form>
</table>
</html>
