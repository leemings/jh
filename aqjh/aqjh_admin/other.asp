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
<html>
<head>
<title>������</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table border="1" cellspacing="1" cellpadding="0" align="center" width="80%" bordercolor="#000000">
<tr>
<td height="9" width="14%">
<div align="center"><a href="binqi.asp"><font color="#0000ff">�ֳֵ���</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="head.asp"><font color="#0000ff">ͷ��</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="body.asp"><font color="#0000ff">����</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="foot.asp"><font color="#0000ff">˫��</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"></div>
<div align="center"><a href="anqi.asp"></a><a href="other.asp"><font color="#0000ff">װ��</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="anqi.asp"><font color="#0000ff">����</font></a></div>
</td>
<td height="19" width="14%" rowspan="2">
<div align="center"><a href="addsu.asp"><font color="#0000ff">����װ��</font></a></div>
</td>
</tr>
<tr>
<td height="9" width="14%">
<div align="center"><a href="dunpai.asp"><font color="#0000ff">�ֳֻ���</font></a></div>
</td>
</tr>
</table>
<p align="center">װ�ι���</p>
<br>
<table border="1" bgcolor="#006699" align="center" width="600" cellpadding="0"
cellspacing="1" bordercolor="#000000">
<tr>
<td height="17" colspan="4" bgcolor="#336633">
<p align="center">�� �� �� ��</p>
<tr>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">װ����</font> </div>
</td>
<td bgcolor="#C4DEFF" height="18">
<div align="center"> <font size="2" color="#000000">�� ��</font> </div>
</td>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">�� ��</font> </div>
</td>
<td height="18" bgcolor="#C4DEFF">
<div align="center"> <font size="2" color="#000000">�� ��</font> </div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM b where b='װ��'"
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
<tr>
<td height="11"><font color="#FFFFFF" size="2"><%=rs("a")%> </font>
<td height="11"><font color="#FFFFFF" size="2">������<%=abs(rs("g"))%> ������<%=abs(rs("f"))%>
</font></td>
<td height="11"><font color="#FFFFFF" size="2"><%=rs("h")%>��</font></td>
<td height="11">
<div align="center"><a href="modifyyao.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF" size="-1">�޸�</font></a>
<font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">ɾ��</font></a></font></div>
</td>
</tr>
<%
rs.movenext
loop
%>
</table>
</body>
</html>