<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
larenseek=Request.Form("larenseek")
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>���˲�ѯ����</title>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"> <font color="#CC0000" face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font>
<br>
��л��������Щ�������������ǽ����ģ�<br>
<table width="650" border="1" cellpadding="0" cellspacing="0" height="13" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="SELECT * FROM �û� where �ȼ�>1 and ������='"& larenseek &"' order by lasttime"
Set Rs=conn.Execute(sql)
%>
<td width="20" height="10">
<div align="center">ID</font></div>
</td>
<td width="47" height="10">
<div align="center">����</font></div>
</td>
<td width="25" height="10">
<div align="center">�Ա�</font></div>
</td>
<td width="63" height="10">
<div align="center">����</font></div>
</td>
<td width="75" height="10">
<div align="center">���</font></div>
</td><td width="75" height="10">
<div align="center">ע��ip</font></div>
</td><td width="75" height="10">
<div align="center">���ip</font></div>
</td>
<td width="75" height="10">
<div align="center">����½ʱ��</font></div>
</td>
<td width="51" height="10">
<div align="center">�����ȼ�</font></div>
</td>
<td width="35" height="10">
<div align="center">��¼</font></div>
</td>
</tr>
<%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
<tr>
<td width="28" height="30">
<div align="center"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="showuser.asp?username=<%=rs("����")%>"><font color="#FF9900"><%=rs("����")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><%=rs("�Ա�")%></font></div>
</td>
<td width="63" height="30">
<div align="center"><%=rs("����")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><%=rs("���")%></font></div>
</td><td width="75" height="30">
<div align="center"><%=rs("regip")%></font></div>
</td><td width="75" height="30">
<div align="center"><%=rs("lastip")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><%=rs("lasttime")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><%=rs("�ȼ�")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><%=rs("times")%></font></div>
</td>
</tr>
<%
rs.movenext
loop
conn.close
%>
</table>
<div align="center"><font color="#000000">��������:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">��</font>
</div>
</body></html>