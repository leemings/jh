<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>属性转换</title>
<link REL="StyleSheet" HREF="../style.css" TITLE="Contemporary">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false >
<p align="center">　</p>
<div align="center">
  <center>
<table border="0" width="225" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100" align="center"><a href="hxie2.asp?id=1"><img border="0" src="../image/05s.gif"></a><br>1000的防体力换成1000的内力</td>
    <td width="108" align="center"></td>
    <td width="107" align="center"><a href="hxie2.asp?id=2"><img border="0" src="../image/07s.gif"></a><br>1000的内力转换成1000的体力</td>
  </tr>
  <tr>
    <td width="100" align="center"></td>
    <td width="108" align="center"><img border="0" src="../image/strut4.gif"></td>
    <td width="107" align="center"></td>
  </tr>
  <tr>
    <td width="100" align="center"></td>
    <td width="108" align="center"><font color="#000000">.</font></td>
    <td width="107" align="center"></td>
  </tr>
  <tr>
    <td width="100" align="center"><a href="hxie2.asp?id=4"><img border="0" src="../image/1414.gif"></a><br>1000的防御转换1000的攻击</td>
    <td width="108" align="center"></td>
    <td width="107" align="center"><a href="hxie2.asp?id=3"><img border="0" src="../image/17s.gif"></a><br>1000的攻击转换成1000的防御</td>
  </tr>
</table>

  </center>
</div>

</body>

</html>
