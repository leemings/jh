<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=aqjh_name
grade=Int(aqjh_grade)
%>
<!--#include file="config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As ���� from �û� where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As ���� from �û� where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As ���� from �û� where CDate(lasttime)<now()-60 and ��Ա�ȼ�=0")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers
%>
<html>
<head>
<title>�ʺŹ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center">�˺Ź��� <br>
</div>
<table width="390" border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="3" align="center">
<tr bgcolor="#0099FF">
<td width="388"><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">�����ʺ�</font></td>
</tr>
<tr>
<td class=p150 width="388" height="34"> �� <a href="manaccdel7.asp?page=1">��������ֻ�ù�һ�ε��ʺ�</a><br>
�� <a href="maingl3.asp">�鿴��������δ��¼�Ļ�Ա</a><br>
�� <a href="maingl.asp">���������û�����ȼ�Ϊ1(����5������Ա����)</a><br>
�� <a href="maingl1.asp">�������е�����żΪ��</a><br>
�� <a href="maingl2.asp">���������ֽ�/����Ϊ���û�Ϊ0</a></td>
</tr>
</table>
<br>
<br>
<table width="390" border="1" cellspacing="0" cellpadding="3" align="center">
  <tr bgcolor="#0099FF"> 
    <td><font face="Wingdings" color="#FFFFFF">1</font><font color="#FFFFFF">��������</font></td>
  </tr>
  <tr> 
    <td height="26">��<a href="gl1.asp">ɾ��10��ǰ,����½3���û���<font color="#0000FF"><b><%=user1%></b>��</font></a> 
      <br>
      ��<a href="gl2.asp">ɾ��10��ǰ�ݵ�allvalueС��5�û���<b><font color="#0000FF"><%=user2%></font></b><font color="#0000FF">��</font></a> 
	<br>
     ��<a href="gl3.asp">ɾ�������´�δ��½���û��У�<font color="#0000FF"><b><%=user3%></b>��</font></a> 
     <br>
    </td>
  </tr>
</table>
</body>
</html>