<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<!--#include file="data.asp"-->
<%
sql="select * from �Խ��� where ӵ����='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "�㻹û�п���!"
  window.close()
</script>
<%
response.end
end if
%>
<html>
<head>
<title><%=aqjh_name%>�Ŀ�ʯһ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "����"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFFFFF"><b><%=session("myname")%>�����п�ʯ���</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td height="18"><font color="#CCCCFF"><b>��ʯ</b></font></td>
    <td height="18"><font color="#CCCCFF"><b>����</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">��</font></b></td>
    <td><font color="#FFFFFF"><%=rs("��")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">��</font></b></td>
    <td><font color="#FFFFFF"><%=rs("��")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">ͭ</font></b></td>
    <td><font color="#FFFFFF"><%=rs("ͭ")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="36"> 
      <input type="submit" name="Submit" value="�ر�" onClick="window.close()">
    </td>
  </tr>
</table>
</body>
</html>
<%
rs.close
       set rs=nothing
       conn.close
       set conn=nothing%>