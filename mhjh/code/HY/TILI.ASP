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
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<link rel=stylesheet href="../style.css">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align="center"><font color="#000000" size="4">会员补给店</font></p>
<div align="center">
  <center>
  <table border="0" width="61%" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%" colspan="2">
        <p align="center"><a href="tili2.asp"><img border="0" src="../image/dtps.gif">
        </a>
        <p align="center">体力店</td>
    </tr>
    <tr>
      <td width="50%" align="center"><a href="tili3.asp"><img border="0" src="../image/0.gif">
        </a>
        <p>内力店</p>
      </td>
      <td width="50%" align="center"><a href="tili2.asp"><img border="0" src="../image/02s.gif">
        </a>
        <p>精力店</p>
      </td>
    </tr>
  </table>
  </center>
</div>
</body>
</html>
