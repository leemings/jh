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
sql="select * from ��ʯ"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>�����Խ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "����"; font-size: 12px}
a:link {  font-family: "����"; font-size: 12px; text-decoration: none}
a:hover {  font-family: "����"; font-size: 12px; color: #FF6666; text-decoration: underline}
a:visited {  font-family: "����"; font-size: 12px; text-decoration: none}
-->
</style>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#376D95">
<table width="72%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="57" width="74%"><b><font color="#FFFF99"><font size="3"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����Խ�</font></font></b> 
      <br>
      <br>
      <font color="#CCCCFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���Խ�ɣ�����һ�ճ�Ϊ������ 
      (��ʯ�����������)</font><font color="#FFFFFF"> </font></td>
    <td height="57" width="26%"><b><a href="#" onclick="MM_openBrWindow('sgou.asp','','width=400,height=200')"><font color="#CCCCFF">��ʯ�չ�</font></a><br>
      <br>
      <a href="#" onclick="MM_openBrWindow('seeme.asp','','width=400,height=200')"><font color="#CCCCFF">���п�ʯ</font></a></b></td>
  </tr>
  <tr align="center"> 
    <td height="197" colspan="2"><img src="worldmap.gif" width="360" height="182" usemap="#Map" border="0"><map name="Map"><area shape="rect" coords="41,22,55,34" href="goto.asp"><area shape="rect" coords="104,21,119,32" href="goto.asp"><area shape="rect" coords="66,55,77,66" href="goto.asp"><area shape="rect" coords="92,58,102,69" href="goto.asp"><area shape="rect" coords="88,79,89,83" href="goto.asp"><area shape="rect" coords="113,145,130,155" href="goto.asp"><area shape="rect" coords="94,82,106,92" href="goto.asp"><area shape="rect" coords="165,30,181,41" href="goto.asp"><area shape="rect" coords="195,43,210,56" href="goto.asp"><area shape="rect" coords="178,83,192,98" href="goto.asp"><area shape="rect" coords="237,64,250,76" href="goto.asp"><area shape="rect" coords="302,95,318,106" href="goto.asp"><area shape="rect" coords="293,22,305,39" href="goto.asp"><area shape="rect" coords="210,113,221,124" href="goto.asp"><area shape="rect" coords="282,73,295,84" href="goto.asp"></map></td>
  </tr>
  <tr align="center"> 
    <td height="57" colspan="2"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999">
        <tr align="center" bgcolor="#376D95"> 
          <td colspan="3" height="22"><font color="#FFFFFF"><b>��ʯ�������</b></font></td>
        </tr>
        <tr align="center" bgcolor="#376D95"> 
          <td height="18"><font color="#FFFFFF"><b>��ʯ</b></font></td>
          <td height="18"><font color="#FFFFFF"><b>�ܵ��</b></font></td>
          <td height="18"><font color="#FFFFFF"><b>������</b></font></td>
        </tr>
        <%
 do while not rs.eof
%> 
        <tr align="center" bgcolor="#376D95"> 
          <td><font color="#FFFFFF"><b><%=rs("����")%></b></font></td>
          <td><font color="#FFFFFF"><%=rs("�ܵ�")%></font></td>
          <td><font color="#FFFFFF"><%=rs("����")%></font></td>
        </tr>
        <%
 rs.movenext
 loop
%> 
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<%
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
%>