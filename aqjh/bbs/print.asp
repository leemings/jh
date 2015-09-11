<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))

sql="Select * From [forum] where ID="&ID&""
rs.Open sql,Conn,1
forumid=rs("forumid")
ReList=rs("ReList")
rs.close

sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
bbsname=rs("bbsname")
moderated=rs("moderated")
pass=rs("pass")
password=rs("password")
userlist=rs("userlist")
rs.close

sql="select * from forum where ID="&id&""
Set Rs=Conn.Execute(sql)

%>
<!-- #include file="inc/validate.asp" -->


<table cellpadding=0 cellspacing=0 border=0 width=97% align=center><tr><td>
- <b><%=clubname%></b> ( http://<%=Request.ServerVariables("server_name")%>/Default.asp )<br>

-- <b><%=bbsname%></b> ( http://<%=Request.ServerVariables("server_name")%>/ShowForum.asp?forumid=<%=forumid%> )<br>

--- <b><%=rs("topic")%></b> ( http://<%=Request.ServerVariables("server_name")%><%=replace(Request.ServerVariables("script_name"),"print.asp","")%>ShowPost.asp?id=<%=id%> )
<br><br>
<hr><p>
作者：<%=rs("username")%><br>
发表时间：<%=rs("posttime")%><br><br>
<%=rs("content")%></p><hr>

<%
replies=rs("replies")
rs.close

if replies > 0 then
sql="select * from [reforum"&ReList&"] where topicid="&id&" order by id"
Set Rs=Conn.Execute(sql)
Do While Not RS.EOF
%><p>
作者：<%=rs("username")%><br>
发表时间：<%=rs("posttime")%><br><br>
<%=rs("content")%></p><hr>
<%
RS.MoveNext
loop
RS.Close
end if

htmlend
%>