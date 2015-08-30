<%
function Checkin(s) 
s=trim(s) 
s=replace(s," ","&amp;nbsp;") 
s=replace(s,"'","&amp;#39;") 
s=replace(s,"""","&amp;quot;") 
s=replace(s,"&lt;","&amp;lt;") 
s=replace(s,"&gt;","&amp;gt;") 
Checkin=s 
end function 
function IsValidEmail(email)
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if
end function
sub error()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Movie" content="flash,flash下载,flash频道,E线,快乐江湖,一线网络,音乐,音乐听吧,视听,回首当年,51eline,happyjh.com,www.happyjh.com,flash.happyjh.com,eline_email@etang.com">
<meta name="Author" content="www.happyjh.com,flash.happyjh.com">
<title>快乐江湖总站</title>
<link rel=stylesheet href=images/style.css>
</head>
<body topmargin="111" leftmargin="0" oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div align="center">
<center>
<table border="0" cellspacing="0" width="60%">
<tr>
<td width="100%" bgcolor="#3F4A5C">
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%" bgcolor="#FFFFFF" height="80" align="center">
<p><b>Error!&nbsp; <%=errmsg%></b>
<BR>
<b><a href="javascript:history.go(-1)">...::: 点 此 返 回 
:::...</a></b></p>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
</center>
</div>
</body> 
</html> 
<%
end sub
%>