<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>���ɹ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<script language="JavaScript">
function shutwin()
{
window.close();
return;
}
</script>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
set rs=server.createobject("adodb.recordset")
conn.open Application("aqjh_usermdb")
rs.open "Select * From p",conn,0,1
if not rs.eof then
%><font color="#FF3333">[ ���ɹ��� ]</font></font></p>
<p align="center">�����ɹ������棬���Ͽ��Ը�����������ɾ��������ݣ����޸��������ϣ�</p>
<table width="700" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
  <tr bgcolor="#336633"> 
    <td align="center" width="57"><font color="#FFFFFF" size="2">&nbsp;�� ��</font></td>
    <td width="40" align="center"><font size="2" color="#FFFFFF">����</font></td>
    <td width="65" align="center"><font color="#FFFFFF" size="2">�� ��</font></td>
    <td width="270" align="center"><font color="#FFFFFF" size="2">���Ƽ����ʽ</font></td>
    <td width="161" align="center"><font size="2" color="#FFFFFF">�� ��</font></td>
    <td width="62" align="center"><font size="2" color="#FFFFFF">����</font></td>
    <td width="66" align="center"><font color="#FFFFFF" size="2">�� ��</font></td>
  </tr>
  <%do while not rs.eof%>
  <tr bgcolor="#FFFFFF"> 
    <td width="57"> 
      <div align="center"><font size="2"><%=rs("a")%></font></div>
    </td>
    <td width="40"> 
      <div align="right"><font size="2"><%=rs("c")%>��</font></div>
    </td>
    <td width="65"> 
      <div align="center"><font size="2"><%=rs("b")%></font></div>
    </td>
    <td width="270"> 
      <div align="center"><font size="2"> <%=rs("e")%><br>
        </font>[<font size="2"><%=rs("g")%></font>]</div>
    </td>
    <td width="161"><font size="2"><%=rs("d")%></font></td>
    <td width="62"> 
      <div align="center"><font size="2"><%=rs("h")%></font></div>
    </td>
    <td width="66"> 
      <div align="center"><font size="2"><%if rs("a")<>"�ٸ�" then%>
<a href=delzm.asp?username=<%=rs("a")%> title="ɾ������">��</a> 
        <a href="delmp.asp?rs=<%=rs("a")%>&action=delete" title="ɾ������">ɾ</a> </font><font size="2"><a href="managemp.asp?id=<%=rs("a")%>" title="�޸��������ϣ�">��</a>
<%else%>
���ܲ���
<%end if%>
</font></div>
    </td>
  </tr>
  <%rs.movenext
loop %>
  <% else %>
</table>
<table width="54%" border="0" cellspacing="0" cellpadding="0" height="40" align="center">
<tr>
<td>
<div align="center">ɾ�����ɻὫ���������и����ɵ���--������Ϊ���ޡ������Ϊ���ޡ���ͬʱɾ����̳�и����ɵ��������ӣ����ϣ����Ǻ��ٰ���ɾ��������<br>
(ɾ��������ʱ�����������˵��κ�״̬������Ҫ����Ҳ�а�)</div>
</td>
</tr>
</table>
<p>&nbsp;</p>
<table align="center" width="197">
<td height="14">
<div align="center">��û��һ�������أ�<% rs.close
set rs=nothing
end if%> </div>
</table>
<div align="center"></div>
<p align="center"><font color="#FF0000">[<a href="adminmpp.asp">====ˢ��====</a>]</font>
</body>