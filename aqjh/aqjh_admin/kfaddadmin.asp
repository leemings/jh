<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title>��Ա���ݿ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>���ݿ�</p>
<table width="572" align="center" border="1" align="center" cellspacing="1" cellpadding="4" bordercolor="#B8AF86">
<tr>
<td>
<p align="center">&nbsp;</p>
<form method="POST" action="kfaddadminok.asp">
<div align="center">
��Ա�ȼ���
<select name="hygrade">
<option value="1" selected>һ����Ա</option>
<option value="2">������Ա</option>
<option value="3">������Ա</option>
<option value="4">�ļ���Ա</option>
<option value="5">�弶��Ա</option>
<option value="6">������Ա</option>
<option value="7">�߼���Ա</option>
<option value="8">�˼���Ա</option>
</select>
<br>
����Ŀ��
<input type="text" name="kfnum" size="10" maxlength="10">
<br>
<input type="submit" value="��ʼ�����Ϳ�" name="B1" class="p9">
</div>
<br>
</div>
<div align="center">
<center>
</center>
</div>
</form>
<p align="center">[վ������ͨ���˴����޸��û�����������]</p>
<p align="center"><a href="../jhmp/hy.asp">���л�Ա�б�</a> <a href="fine.asp">�����û�����</a></p>
</td>
</tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>
