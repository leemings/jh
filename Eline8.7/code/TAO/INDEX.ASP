<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<!--#include file="data.asp"-->
<%
sql="select * from 矿石"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>江湖淘金♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
a:link {  font-family: "宋体"; font-size: 12px; text-decoration: none}
a:hover {  font-family: "宋体"; font-size: 12px; color: #FF6666; text-decoration: underline}
a:visited {  font-family: "宋体"; font-size: 12px; text-decoration: none}
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
    <td height="57" width="74%"><b><font color="#FFFF99"><font size="3"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;江湖淘金</font></font></b> 
      <br>
      <br>
      <font color="#CCCCFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;来淘金吧，包你一日成为百万富翁 
      (矿石就在世界各地)</font><font color="#FFFFFF"> </font></td>
    <td height="57" width="26%"><b><a href="#" onclick="MM_openBrWindow('sgou.asp','','width=400,height=200')"><font color="#CCCCFF">矿石收购</font></a><br>
      <br>
      <a href="#" onclick="MM_openBrWindow('seeme.asp','','width=400,height=200')"><font color="#CCCCFF">我有矿石</font></a></b></td>
  </tr>
  <tr align="center"> 
    <td height="197" colspan="2"><img src="worldmap.gif" width="360" height="182" usemap="#Map" border="0"><map name="Map"><area shape="rect" coords="41,22,55,34" href="goto.asp"><area shape="rect" coords="104,21,119,32" href="goto.asp"><area shape="rect" coords="66,55,77,66" href="goto.asp"><area shape="rect" coords="92,58,102,69" href="goto.asp"><area shape="rect" coords="88,79,89,83" href="goto.asp"><area shape="rect" coords="113,145,130,155" href="goto.asp"><area shape="rect" coords="94,82,106,92" href="goto.asp"><area shape="rect" coords="165,30,181,41" href="goto.asp"><area shape="rect" coords="195,43,210,56" href="goto.asp"><area shape="rect" coords="178,83,192,98" href="goto.asp"><area shape="rect" coords="237,64,250,76" href="goto.asp"><area shape="rect" coords="302,95,318,106" href="goto.asp"><area shape="rect" coords="293,22,305,39" href="goto.asp"><area shape="rect" coords="210,113,221,124" href="goto.asp"><area shape="rect" coords="282,73,295,84" href="goto.asp"></map></td>
  </tr>
  <tr align="center"> 
    <td height="57" colspan="2"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999">
        <tr align="center" bgcolor="#376D95"> 
          <td colspan="3" height="22"><font color="#FFFFFF"><b>矿石开采情况</b></font></td>
        </tr>
        <tr align="center" bgcolor="#376D95"> 
          <td height="18"><font color="#FFFFFF"><b>矿石</b></font></td>
          <td height="18"><font color="#FFFFFF"><b>总点击</b></font></td>
          <td height="18"><font color="#FFFFFF"><b>总流量</b></font></td>
        </tr>
        <%
 do while not rs.eof
%> 
        <tr align="center" bgcolor="#376D95"> 
          <td><font color="#FFFFFF"><b><%=rs("矿种")%></b></font></td>
          <td><font color="#FFFFFF"><%=rs("总点")%></font></td>
          <td><font color="#FFFFFF"><%=rs("总流")%></font></td>
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