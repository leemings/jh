
<form action=2.asp method=post>
请输入用户名:<input type=txt size=20 name=name><input type=submit value='确认'>

<%
Application("sjjh_user")=Request.form("name")
temp="sjjh_"&"soft"&"jm"
temp1="sjjh_"&"user"
user=Application(temp1)
usersn=Application(temp)
for ii=1 to len(user)
ll=ll+asc(mid(user,ii,1))
i=i+i+asc(mid(user,ii,1))
dd=asc(mid(user,ii,1))-dd
bb=mid(user,ii,1)
cc=asc(bb)
next 
if ll<>usersn then
%>
序列号:<%=ll%>
<%
end if
%>