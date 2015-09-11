<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))


if id="0" then

if Request("forumid")="0" then
sql="select * from online where username<>'' and eremite<>1"
else
sql="select * from online where forumid="&int(Request("forumid"))&" and username<>'' and eremite<>1"
end if

Set Rs=Conn.Execute(sql)

do while not rs.eof

if NO_count < 6 then
NO_count=NO_count+1
else
NO_count=1
end if


allline=""&allline&"<td width=16% style=word-break:break-all><img src=images/membercode/"&rs("userface")&".gif> <a href=Profile.asp?username="&rs("username")&">"&rs("username")&"</a></td>"

if NO_count = 6 then allline=""&allline&"</tr><tr>"

rs.movenext
loop
rs.close

%>
<SCRIPT>
parent.followImg0.loaded="yes";
parent.followTd0.innerHTML="<TABLE cellSpacing=3 cellPadding=0 width=100% border=0><TR><%=allline%></TR></TABLE>";
</SCRIPT>
<%
else


sql="Select * From forum where id="&id&""
Set Rs=Conn.Execute(sql)
content=rs("content")
forumid=rs("forumid")
rs.close

sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
pass=rs("pass")
password=rs("password")
userlist=rs("userlist")
rs.close


if pass="0" then
content="该论坛暂时关闭，不再接受访问！"
elseif pass="3" then
if instr("|"&userlist&"|","|"&Request.Cookies("username")&"|")<=0 then content="该论坛并未授权您访问！"
end if

if password<>empty then
if password<>Request.Cookies("password") then content="该论坛需要密码才能访问！"
end if

%>
<SCRIPT>
parent.followImg<%=id%>.loaded="yes";
parent.followTd<%=id%>.innerHTML='<table width=100% cellpadding=10 style=TABLE-LAYOUT:fixed><tr><td width=100% style=WORD-WRAP:break-word><%=content%></td></tr></table>';
</SCRIPT>
<%
end if
responseend
%>