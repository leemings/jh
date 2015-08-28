<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>温泉浴中心</title>
<link rel="stylesheet" href="../style.css" type="text/css">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<div align="center">
<center>
<table border="0" width="479" cellspacing="0" cellpadding="0">
<tr>
<td width="847">
<form method="POST" action="CHECKSEX.ASP">
<p align="center"><b><font face="宋体" size="4">魔幻江湖嘉宾俱乐部</font></b></p>
<p align="center"><font size="2" face="宋体">欢迎光临，我们有最好的按摩小姐为你服务，请你慢慢享受温泉浴</font></p>
<table width="100%" cellspacing="0" height="126"  border="1" bordercolordark="#FFFFFF" bordercolorlight="#000000">
<tr bgcolor="#33FF33">
<td class="p2" width="100%" bgcolor="#F0BB85" height="18">
<p align="center"><font size="2" face="宋体">欢 迎 进 入 温 泉 洗   
浴 中 心&nbsp;</font>   
</td>   
</tr>   
<tr bgcolor="#FFFFFF">   
<td class="p3" width="100%" height="23">   
<p align="center"><font size="2" face="宋体">洗浴中心：门票50000元&nbsp;   
加体力200</font></p>   
</td>   
</tr>   
<tr bgcolor="#FFFFFF">   
<td class="p2" width="100%" height="34">   
<p align="center"><font size="2" face="宋体"><input type="radio" value="男"   
name="SEX" checked>      男宾部&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
<input type="radio" value="女" name="SEX">    女宾部</font></p>   
</td>   
</tr>   
<tr bgcolor="#FFFFFF">   
<td class="p3" width="100%" height="27">   
<p align="center"><input type="submit" value="进入洗浴"   
name="B1"></p>   
</td>   
</tr>   
</table>   
</form>   
</td>   
<td width="132">   
<img border="0" src="../image/414ed.gif" width="69" height="320">   
</td>   
</tr>   
</table>   
</center>   
</div>   
<center><br>   
</center>   
</body>   
</html>   
   
