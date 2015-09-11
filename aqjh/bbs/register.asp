<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if CloseRegUser = 1 then error("<li>±¾ÂÛÌ³ÔİÊ±²»¿ª·ÅĞÂÓÃ»§×¢²á£¡")
username=HTMLEncode(Trim(Request("username")))
errorchar=array(" ","¡¡","	","","","","#","`","|","%","&","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error2("ÓÃ»§ÃûÖĞ²»ÄÜº¬ÓĞÌØÊâ·ûºÅ")
next
if Request("menu")="face" then
%>
<body topmargin=0>
<title>BBSxp--Í·ÏñÁĞ±í - Powered By BBSxp</title>
<table border=0 width=100% cellspacing=1 cellpadding=1>
<tr class=a1><td colspan="10" height="25" align=center>BBSxpÍ·Ïñ</td></tr>
<tr align=center>
<script>
var ii=1
for (var i=1; i <= 84; i++) {
ii++
document.write("<td class=a3><img src=images/face/"+i+".gif><br>"+i+"</td>");
if(ii >5){document.write("</tr><tr align=center>");ii=1}
}
</script>
</tr>
</table>
<%
responseend
elseif Request("menu")="Check" then
If conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then
response.write "ÓÃ»§Ãû&quot; <font color=red>"&HTMLEncode(username)&"</font> &quot;¿ÉÒÔÕı³£×¢²á£¡"
else
response.write "ÄúËùÑ¡µÄÓÃ»§Ãû&quot; <font color=red>"&username&"</font> &quot;ÒÑ¾­ÓĞÓÃ»§Ê¹ÓÃ£¬ÇëÁíÍâÑ¡ÔñÒ»¸öÓÃ»§Ãû¡£"
end if
responseend
end if

top
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then
error("<li>ÌáÊ¾<li>½­ºşÂÛÌ³½ûÖ¹×¢²á")
elseif Request("menu")="write" then
%>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> ¡ú ×¢²áĞÂÓÃ»§</td>
</tr>
</table><br>
<SCRIPT>valigntop()</SCRIPT>
<table width=97% align=center cellspacing=0 cellpadding=0 border=0 class=a1>
<form method="POST" name="form" onsubmit="return VerifyInput();">
<input type=hidden name=sessionid value=<%=session.sessionid%>>
<tr>
<td>
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr bgcolor="FFFFFF" class=a1>
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;¸öÈËÉçÇøĞÅÏ¢</b>£¨ÒÔÏÂÄÚÈİ±ØÌî£©</td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="46%"><b>³ö´íĞÅÏ¢£º</b></td>
<td class=a3 height="5" align="left" valign="middle" width="54%">
&nbsp;½­ºşÂÛÌ³½ûÖ¹×¢²á
</td>
</tr>
</table>
</td>
</tr></table>

<SCRIPT>valignbottom()</SCRIPT>

<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="ÖĞ¹ú"}
else
if (language=="zh-hk"){country="ÖĞ¹úÏã¸Û"}
else
if (language=="zh-tw"){country="ÖĞ¹úÌ¨Íå"}
else
if (language=="zh-sg"){country="ĞÂ¼ÓÆÂ"}
else
country=""
document.form.country.value=country



function Check(){var Name=document.form.username.value;
window.open("register.asp?menu=Check&username="+Name,"","width=200,height=20");
}


function VerifyInput()
{
username=document.form.username.value
//if (/[^\x00-\xff]/g.test(username)){alert("ÓÃ»§Ãû²»ÄÜº¬ÓĞºº×Ö");return false;}

if (username == "")
{
alert("ÇëÊäÈëÄúµÄÓÃ»§Ãû");
document.form.username.focus();
return false;
}

var mail = document.form.usermail.value;
if(mail.indexOf('@',0) == -1 || mail.indexOf('.',0) == -1){
alert("ÄúÊäÈëµÄEmailÓĞ´íÎó\nÇëÖØĞÂ¼ì²éÄúµÄEmail");
document.form.usermail.focus();
return false;
}

if (document.form.usermail.value == "")
{
alert("ÇëÊäÈëÄúµÄEMAILµØÖ·");
document.form.usermail.focus();
return false;
}

return true;
}
</SCRIPT>
<%
elseif Request("menu")="" then
%><center>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> ¡ú ×¢²áĞÂÓÃ»§</td></tr></tbody></table>
<br>
<textarea name="textarea" rows="18" readOnly cols="90"><%=Conn.Execute("Select agreement From [clubconfig]")(0)%></textarea><br><br>
<input type="submit" value=" Í¬ Òâ " onclick="window.location.href='register.asp?menu=write';">
<input type="submit" value=" ²»Í¬Òâ " onclick="history.back()">
<%
end if


htmlend
%>