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
sql="select * from 淘金者 where 拥有者='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "你还没有矿呢!"
  window.close()
</script>
<%
response.end
end if
%>
<html>
<head>
<title><%=aqjh_name%>的矿石一览</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFFFFF"><b><%=session("myname")%>你现有矿石情况</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td height="18"><font color="#CCCCFF"><b>矿石</b></font></td>
    <td height="18"><font color="#CCCCFF"><b>数量</b></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">金</font></b></td>
    <td><font color="#FFFFFF"><%=rs("金")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">银</font></b></td>
    <td><font color="#FFFFFF"><%=rs("银")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td><b><font color="#FFFFFF">铜</font></b></td>
    <td><font color="#FFFFFF"><%=rs("铜")%></font></td>
  </tr>
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="36"> 
      <input type="submit" name="Submit" value="关闭" onClick="window.close()">
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