<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='会员'",conn,2,2
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script>
window.moveTo(30,30);
</script>
<style type="text/css">
A {COLOR: #ffffff; TEXT-DECORATION: none; TEXT-TRANSFORM: none}
A:hover {COLOR:#C46200; TEXT-DECORATION: underline}
BODY {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
</style>
<title><%=Application("sjjh_chatroomname")%>会员办法</title>
</head>
<body topmargin="0" leftmargin="0" bgcolor="#336699" text="#FFFFFF">
<table border="0" width="375" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" height="6">
</td>
</tr>
<tr>
<td width="100%">
      <p align="center"><span style="FONT-SIZE: 14px"><b><font color="#FFFF00"><%=Application("sjjh_chatroomname")%>会员简介</font></b></span>
    </td>
</tr>
</table>
<table border="0" width="484" cellspacing="0" cellpadding="0" height="427" align="center">
  <tr> 
    <td width="1" height="262">&nbsp;</td>
    <td width="585" height="262"><%=rs("c")%><%=rs("d")%></td>
  </tr>
  <tr> 
    <td width="1" height="22">&nbsp;</td>
    <td width="585" height="22"> 
      <p align="center"><a href="javascript:window.close()"><font color="#FFFFFF"><br>
        关闭窗口</font></a> 
    </td>
  </tr>
</table>
<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
</body>
</html>