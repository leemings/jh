<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%
id=request("id")
if instr(id,"��")<>0 then
		Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
		Response.End
end if
my=request("my")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css><!--
.text {  font-size: 12px; color: #FFFFFF; text-decoration: none}
-->
</style>
<link rel="stylesheet" href="../chat/READONLY/Style.css">
<title>���ɾ�������wWw.51eline.com��</title>
</head>
<body leftmargin="0" topmargin="5" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" background="../chat/lvbgcolor.gif" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<div align="center">
<b><font size="+2" color="#008080">���ɾ�����</font></b><br>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ���� FROM �û� where ����='"&sjjh_name&"'",conn

gongsi=rs("����")

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM �û� where ����='"&gongsi&"' and ���<>'����' and ���<>'����' and ���<>'����' and ���<>'����' order by ���ɻ��� DESC",conn
%>
</div>

<table border="1" width="80%" cellspacing="0" cellpadding="3" height="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#FFFFFF">
  <tr align="center"> 
    <td height="20" colspan="6"> 
      <table border="1" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFCCCC" align="center">
        <tr> 
          <td width="25%" valign="middle" align="center"><a href="bentop4.asp"><font color="#0000ff">���ŵ�������</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop1.asp"><font color="#ff0000">������������</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop2.asp"><font color="#0000FF">���Ż�������</font></a></td>
          <td width="25%" valign="middle" align="center"><a href="bentop3.asp"><font color="#0000FF">���ų�������</font></a></td>
        </tr>
      
	 </table>
<tr align="center"> 
    <td height="16" width="18%" bgcolor="#669900"><font color="#FFFFFF">����</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">����</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">����</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">����</font></td>
    <td height="16" width="16%" bgcolor="#669900"><font color="#FFFFFF">ս���ȼ�</font></td>
    <td height="16" width="18%" bgcolor="#669900"><font color="#FFFFFF">���ɻ���</font></td>

<%do while not rs.bof and not rs.eof%>
  <tr align="center"> 
    <td height="19"><font size="2"><%=rs("����")%></font></td>
    <td height="19"><font size="2"><%=rs("����")%></font></td>
    <td height="19"><font size="2"><%=rs("����")%></font></td>
    <td height="19"><font size="2"><%=rs("mvalue")%></font></td>
    <td height="19"><font size="2"><%=rs("�ȼ�")%></font></td>
    <td height="19"><font size="2"><%=rs("���ɻ���")%></font></td>
<tr>

	<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>

<table border="1" width="80%" cellspacing="0" cellpadding="3" height="0" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#FFFFFF">
  <tr bgcolor="#669900"> 
    <td colspan="3"><b><font color="#FFFFFF">�����������</font></b></td>
  </tr>
  <tr> 
    <td rowspan="2" width="19%" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">����(8��)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>��������</b>��<b> 
      <font color="#FF9900"><font color="#DD2222"> </font></font></b></font><b> 
      <font color="#FFFF00"> 
      ս���ȼ�</font></b><font color="#FFFFFF">25�����ϣ�</font><b><font color="#FFFF00">���ɻ���</font></b><font color="#FFFFFF">30000000������</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"><b><font color="#FFFFFF">ӵ��Ȩ��</font></b><font color="#FFFFFF">��<b> 
      </b>�����̷������ɷ���������������<b>��</b>��ѵ���ӣ����յ���</font></td> 
    <td width="13%" class="text" bgcolor="#008080"><b><font color="#FFFFFF"><a href="zongcai.asp" target="_blank"></a></font><a href="zongcai.asp" target="_blank"><font color="#FF0000">��Ҫ����</font></a><font color="#FFFFFF"></font></b></td>
  </tr>
  <tr> 
    <td rowspan="2" width="19%" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">����(10��)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>��������</b>�� 
      </font> <b><font color="#FFFF00">ս���ȼ�</font></b><font color="#FFFFFF">20���ϣ�</font><b><font color="#FFFF00">���ɻ���</font></b><font color="#FFFFFF">20000000������</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"> 
      <p><b><font color="#FFFFFF">ӵ��Ȩ��</font></b><font color="#FFFFFF">�� �����̷������ɷ����������</font></p> 
    </td>
    <td width="13%" class="text" bgcolor="#008080"><b><font color="#FFFFFF"><a href="mishu.asp" target="_blank"></a></font><a href="mishu.asp" target="_blank"><font color="#FF0000">��Ҫ����</font></a><font color="#FFFFFF"></font></b></td>
  </tr>
  <tr> 
    <td width="19%" rowspan="2" class="text" bgcolor="#008080"><b><font size="2" color="#FFFF00">����(16��)</font></b></td>
    <td colspan="2" class="text" bgcolor="#008080"><font color="#FFFFFF"><b>��������</b>�� 
      </font> <b><font color="#FFFF00">ս���ȼ�</font></b><font color="#FFFFFF">15�����ϣ�</font><b><font color="#FFFF00">���ɻ���</font></b><font color="#FFFFFF">15000000������</font></td>
  </tr>
  <tr> 
    <td width="68%" class="text" bgcolor="#008080"><b><font color="#FFFFFF">ӵ��Ȩ��</font></b><font color="#FFFFFF">�� 
      �����̷������ɷ���</font></td>
    <td width="13%" class="text" bgcolor="#008080"> 
      <div align="left"><b><font color="#FFFFFF"><a href="jingli.asp" target="_blank"></a></font><a href="jingli.asp" target="_blank"><font color="#FF0000">��Ҫ����</font></a><font color="#FFFFFF"></font></b></div>
    </td>
  </tr>
</table>
    <p align="center"><font color="#3366FF">��E�߽�����&#8482;</font></p> 


