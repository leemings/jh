<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�����������</TITLE>
<link rel="stylesheet" href="../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<!--#include file="data.asp"-->
<center>
<font color="red"><b>�������������</b></font><BR><br>
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from ��Ȧ����"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
%>
<font color="#0000FF">������ʼʱ�䣺<font color=red><%=rs("������ʼ")%></font> <BR>��������ʱ�䣺<font color=red><%=rs("��������")%></font></font>
<%if CDate(rs("������ʼ"))>now() then
   	Response.Write "<br><br><b>[��ʾ]</b>������û�ڿ�ʼ�أ�"
	Response.End
  elseif CDate(rs("��������"))<now() then
     	Response.Write "<br><br><b>[��ʾ]</b>�����Ѿ������ˣ���һ�������ɣ�"
	Response.End
  else
%>
<table border=1 bgcolor="#669900" align=center width=350 cellpadding="10" cellspacing="13">
<tr>
<td bgcolor=#FFFFFF align="center"> 
<table bgcolor="#FFFFFF" bordercolor="#FFFF00" border="0">
<tr> 
<td bgcolor="#FFFFFF" height="39">&nbsp;</td>
<td bgcolor="#FFFFFF" height="39" align="center"> 
<form method=POST action='modify.asp'>
<table border="0" cellpadding="0">
<tr> 
<td align="center">�������⣺</td>
<td align="center"> 
<input type=text name=hyname size=12 maxlength="11" value=""> <font size=2 color=red>�����ʲô����</font>
</td>
</tr>
<tr> 
<td align="center">����˵����</td>
<td align="center"> 
<input type=text name=hyTitle size=26 maxlength="32" value="">
</td>
</tr>
</table><br>
<input type=submit value=���� name="submit">
<input type=button value=�ر� onClick="window.close()" name="button">
</form>
</td>
</tr>
</table>
</table>
<%end if
rs.close
end if
%>
<font color=blue>������˵���в�Ҫʹ�ÿո�</font>
</center> 
</body></html>