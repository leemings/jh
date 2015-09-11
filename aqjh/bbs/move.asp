<!-- #include file="setup.asp" -->
<%

id=int(Request("id"))

forumid=Conn.Execute("Select forumid From forum where id="&id)(0)

sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
bbsname=rs("bbsname")
logo=rs("logo")
followid=rs("followid")
rs.close


top
sql="select * from forum where ID="&ID&" and forumid="&forumid&""
Set Rs=Conn.Execute(sql)
%>
<script>
if ("<%=logo%>"!=''){logo.innerHTML="<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="97%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?forumid=<%=forumid%>><%=bbsname%></a> → <a href="ShowPost.asp?id=<%=id%>"><%=rs("topic")%></a> → 移动帖子</td>
		</tr>
	</table><br>

<form method="post" action=manage.asp>
<input type="hidden" value="move" name=menu>
<input type="hidden" value="<%=id%>" name="id">

<table width="50%" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr><td class=a1 height="25"><div align="center"><span>
移动帖子</span></div>
</td></tr><tr><td height="19" valign="top" class=a4>
<div align="center"><p><br>
	<select name=moveid>




<option selected value="">将主题移动到...</option>

<%
rs.close
%>
<%BBSList(0)%>

</select>&nbsp;&nbsp;

<input type=submit value=" 确 定 "> <br>
　</div></td></tr></table>
</form>





<%
htmlend
%>