<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录江湖！"
window.close()
</script>
<%
else
%>
<html>
<head>
<title>大内密探</title>
<link rel=stylesheet href='../../style.css'>
</head>
<body background='../../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false topmargin="50">
<p align="center"><font face="隶书" size="4"><b>大内密探零零发</b></font></p>
<div align="center">
<center>
<table border="0" width="64%" cellspacing="0" cellpadding="0">
<tr>
<td width="100%">我是江湖中最有名气的密探零零发，只要你付得起银两，你就可调查剑侠中你的朋友，恋人，仇敌，妻子....等人的详细资料了。不要偷看人家的私隐阿！不过你放心，没有人知道你做过什么的！慢慢观赏。不要以为美女出浴阿！最多是天仙出浴，小心你的眼睛眼！呵呵！</td>
</tr>
</table>
</center>
</div>
<form method="POST" action="showuser.asp">
<div align="center">
<center>
<p>请在这里输入你要调查对象的名字 <font color="#FFFFFF">
<input type="text" name="search" size="10"
maxlength="10">
<input type="submit" value="调查" name="B1" class="p9"> 
<input 
type="reset" name="Submit" value="清空"> 
</font> 
</p> 
</center> 
</div> 
</form> 
<div align="center"><font color="#FFFFFF"><br> 
</font> 
每次收取调查费1000两<font color="#FFFFFF"> 
<br><br><br> 
</font> 
</div> 
</body> 
</html> 
<% end if %> 
 
 
 
 
 
 
