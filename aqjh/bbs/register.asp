<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if CloseRegUser = 1 then error("<li>本论坛暂时不开放新用户注册！")
username=HTMLEncode(Trim(Request("username")))
errorchar=array(" ","　","	","","","","#","`","|","%","&","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error2("用户名中不能含有特殊符号")
next
if Request("menu")="face" then
%>
<body topmargin=0>
<title>BBSxp--头像列表 - Powered By BBSxp</title>
<table border=0 width=100% cellspacing=1 cellpadding=1>
<tr class=a1><td colspan="10" height="25" align=center>BBSxp头像</td></tr>
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
response.write "用户名&quot; <font color=red>"&HTMLEncode(username)&"</font> &quot;可以正常注册！"
else
response.write "您所选的用户名&quot; <font color=red>"&username&"</font> &quot;已经有用户使用，请另外选择一个用户名。"
end if
responseend
end if

top
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then
error("<li>提示<li>江湖论坛禁止注册")
elseif Request("menu")="write" then
%>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 注册新用户</td>
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
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;个人社区信息</b>（以下内容必填）</td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="46%"><b>出错信息：</b></td>
<td class=a3 height="5" align="left" valign="middle" width="54%">
&nbsp;江湖论坛禁止注册
</td>
</tr>
</table>
</td>
</tr></table>

<SCRIPT>valignbottom()</SCRIPT>

<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="中国"}
else
if (language=="zh-hk"){country="中国香港"}
else
if (language=="zh-tw"){country="中国台湾"}
else
if (language=="zh-sg"){country="新加坡"}
else
country=""
document.form.country.value=country



function Check(){var Name=document.form.username.value;
window.open("register.asp?menu=Check&username="+Name,"","width=200,height=20");
}


function VerifyInput()
{
username=document.form.username.value
//if (/[^\x00-\xff]/g.test(username)){alert("用户名不能含有汉字");return false;}

if (username == "")
{
alert("请输入您的用户名");
document.form.username.focus();
return false;
}

var mail = document.form.usermail.value;
if(mail.indexOf('@',0) == -1 || mail.indexOf('.',0) == -1){
alert("您输入的Email有错误\n请重新检查您的Email");
document.form.usermail.focus();
return false;
}

if (document.form.usermail.value == "")
{
alert("请输入您的EMAIL地址");
document.form.usermail.focus();
return false;
}

return true;
}
</SCRIPT>
<%
elseif Request("menu")="" then
%><center>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → 注册新用户</td></tr></tbody></table>
<br>
<textarea name="textarea" rows="18" readOnly cols="90"><%=Conn.Execute("Select agreement From [clubconfig]")(0)%></textarea><br><br>
<input type="submit" value=" 同 意 " onclick="window.location.href='register.asp?menu=write';">
<input type="submit" value=" 不同意 " onclick="history.back()">
<%
end if


htmlend
%>