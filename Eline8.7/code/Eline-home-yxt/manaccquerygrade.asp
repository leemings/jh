<%@ LANGUAGE=VBScript codepage ="936" %>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(5)<>"�ٸ�" and info(2)<10 or info(0)<>Application("sjjh_admin") then Response.Redirect "manerr.asp?id=242"
%><html>
<head>
<title>�ʺŹ����wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" class=p150 background="../jhimg/bk_hc3w.gif">
<div align="center">
<h1><font color="0099FF">���ʺŹ���</font></h1>
<font color="#FF0000">���������г������ʺš�</font></div>
<div align="center"><a href="txt.asp">����</a> </div>
<hr noshade size="1" color=009900>
<b>��ע�������</b><br>
��ѡ��Ҫ��ѯ�ļ���
<hr noshade size="1" color=009900>
<div align=center>
<table border="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" bgcolor="#E0E0E0" cellpadding="8">
<form method="post" action="manaccquerygradeok.asp">
<tr align="center">
<td>
<table width="100%" border="0" cellpadding="4">
<tr>
<td>Ҫ��ѯ�ļ���</td>
<td>
<select name="grade">
<option value="0">0 ������</option>
<option value="1" selected>1 ��</option>
<option value="2">2 ��</option>
<option value="3">3 ��</option>
<option value="4">4 ��</option>
<option value="5">5 ��</option>
<option value="6">6 ��</option>
<option value="7">7 ��</option>
<option value="8">8 ��</option>
<option value="9">9 ��</option>
<option value="10">10 ��</option>
</select>
</td>
</tr>
<tr>
<td colspan="2" align="center">
<input type="submit" name="Submit" value="��ѯ">
<input type="reset" value="��д" name="reset">
<input type="button" value="����" onClick="javascript:history.go(-1)" name="button">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</div>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "���кţ�<font color=blue>" & Application("sjjh_sn") & "</font>����Ȩ����<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>
