<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
if CloseRegUser = 1 then error("<li>����̳��ʱ���������û�ע�ᣡ")
username=HTMLEncode(Trim(Request("username")))
errorchar=array(" ","��","	","","","","#","`","|","%","&","+",";")
for i=0 to ubound(errorchar)
if instr(username,errorchar(i))>0 then error2("�û����в��ܺ����������")
next
if Request("menu")="face" then
%>
<body topmargin=0>
<title>BBSxp--ͷ���б� - Powered By BBSxp</title>
<table border=0 width=100% cellspacing=1 cellpadding=1>
<tr class=a1><td colspan="10" height="25" align=center>BBSxpͷ��</td></tr>
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
response.write "�û���&quot; <font color=red>"&HTMLEncode(username)&"</font> &quot;��������ע�ᣡ"
else
response.write "����ѡ���û���&quot; <font color=red>"&username&"</font> &quot;�Ѿ����û�ʹ�ã�������ѡ��һ���û�����"
end if
responseend
end if

top
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then
error("<li>��ʾ<li>������̳��ֹע��")
elseif Request("menu")="write" then
%>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ע�����û�</td>
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
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;����������Ϣ</b>���������ݱ��</td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="46%"><b>������Ϣ��</b></td>
<td class=a3 height="5" align="left" valign="middle" width="54%">
&nbsp;������̳��ֹע��
</td>
</tr>
</table>
</td>
</tr></table>

<SCRIPT>valignbottom()</SCRIPT>

<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="�й�"}
else
if (language=="zh-hk"){country="�й����"}
else
if (language=="zh-tw"){country="�й�̨��"}
else
if (language=="zh-sg"){country="�¼���"}
else
country=""
document.form.country.value=country



function Check(){var Name=document.form.username.value;
window.open("register.asp?menu=Check&username="+Name,"","width=200,height=20");
}


function VerifyInput()
{
username=document.form.username.value
//if (/[^\x00-\xff]/g.test(username)){alert("�û������ܺ��к���");return false;}

if (username == "")
{
alert("�����������û���");
document.form.username.focus();
return false;
}

var mail = document.form.usermail.value;
if(mail.indexOf('@',0) == -1 || mail.indexOf('.',0) == -1){
alert("�������Email�д���\n�����¼������Email");
document.form.usermail.focus();
return false;
}

if (document.form.usermail.value == "")
{
alert("����������EMAIL��ַ");
document.form.usermail.focus();
return false;
}

return true;
}
</SCRIPT>
<%
elseif Request("menu")="" then
%><center>
<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2><tbody><tr class=a3><td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> �� ע�����û�</td></tr></tbody></table>
<br>
<textarea name="textarea" rows="18" readOnly cols="90"><%=Conn.Execute("Select agreement From [clubconfig]")(0)%></textarea><br><br>
<input type="submit" value=" ͬ �� " onclick="window.location.href='register.asp?menu=write';">
<input type="submit" value=" ��ͬ�� " onclick="history.back()">
<%
end if


htmlend
%>