<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then%>
<script language=vbscript>
MsgBox "�Բ����㻹û�е�¼������"
window.close()
</script>
<%else
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=Application("sjjh_chatroomname")%>����Ȫԡ���ġ�wWw.51eline.com��</title>
<link rel="StyleSheet" href="../../css.css" title="Contemporary">
</head>

<body leftmargin="0" topmargin="0">
<div align="center">
<center>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr bgcolor="#FFFFFF">
<td width="86"></td>
<td width="33"><img src="../../images/yc4.gif" width="84" height="193"></td>
        <td width="140" valign="bottom">&nbsp;</td>
<td width="519" valign="bottom">&nbsp;</td>
</tr>
<tr bgcolor="#000000">
<td width="86"></td>
<td width="33"><img src="../../images/yc3.jpg" width="84" height="242"></td>
        <td width="140" valign="top">&nbsp;</td>
<td width="519" valign="top">
<form method="POST" action="CHECKSEX.ASP">
<br>
<br>
<table width="100%" cellspacing="1">
<tr>
                <td class="p2" width="100%"> <font color="#CCCCCC">�� ӭ �� �� E �� 
                  �� �� �� Ȫ ϴ ԡ �� ��&nbsp;</font> </td>
</tr>
<tr >
                <td class="p3" width="100%" height="29"> <font color="#FFFFFF">����ϴԡ��Ͳ����칤�ʣ���</font>����1000 
                </td>
</tr>
<tr>
<td class="p2" width="100%" height="50">

<input type="radio" value="��"
name="SEX" checked>
<font color="#FFFFFF"> <font color="#CCCCCC">�б���&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="radio" value="Ů" name="SEX">
Ů����</font></font> </td>
</tr>
<tr>
<td class="p3" width="100%">


<p>&nbsp;</p>
<p>
<input type="submit" value="����ϴԡ"
name="B1">
</p>
</td>
</tr>
</table>
</form>
<%=Application("sjjh_chatroomname")%>��Ȩ���� ���й������޸�</td>
</tr>
</table>
</center>
</div>

</body>

</html>
<%
end if
%>

<html><script language="JavaScript">                                                                  </script></html>


<html><script language="JavaScript">                                                                  </script></html>


<html></html>

<html></html>
<html></html>
<html></html>
<html></html>
<html></html>
<html></html>


