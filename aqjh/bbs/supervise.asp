<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

forumid=int(Request("forumid"))


moderated=Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0)
if ""&moderated&""="" then moderated="|"
moderated=split(moderated,"|")


if membercode<4 and moderated(0)<>Request.Cookies("username") then
error("<li>您的权限不够<li>只有正版主 <font color=red>"&moderated(0)&"</font> 、管理员、社区区长才拥有此权限")
end if


select case Request("menu")
case "up"

if Request("bbsname")="" then error("<li>请输入论坛名称")

if Len(Request("intro"))>255 then  error("<li>论坛简介不能大于 255 个字节")



master=split(HTMLEncode(Request("moderated")),"|")
for i = 0 to ubound(master)
If conn.Execute("Select id From [user] where username='"&master(i)&"'" ).eof and master(i)<>"" Then error("<li>"&master(i)&"的用户资料还未注册")
next


sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn,1,3
rs("bbsname")=HTMLEncode(Request("bbsname"))
rs("moderated")=Request("moderated")
rs("intro")=HTMLEncode(Request("intro"))
rs("icon")=HTMLEncode(Request("icon"))
rs("logo")=HTMLEncode(Request("logo"))
rs.update

rs.close

log("更新论坛（ID:"&forumid&"）的信息！")

message="<li>更新成功！<li><a href=ShowForum.asp?forumid="&forumid&">返回论坛</a><li><a href=Default.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")


case "delusertopic"

username=HTMLEncode(Request("username"))
if username="" then error("<li>您没有输入字符！")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where username='"&username&"' and forumid="&forumid&" ")
error2("已经将 "&username&" 的帖子全部删除了！")

case "deltopic"


topic=HTMLEncode(Request("topic"))
if topic="" then error("<li>您没有输入字符！")
conn.execute("update [forum] set toptopic=0,deltopic=1,lastname='"&Request.Cookies("username")&"',lasttime='"&now()&"' where topic like '%"&topic&"%' and forumid="&forumid&" ")
error2("已经将标题里包含有 "&topic&" 的帖子全部删除了！")


end select




sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)


%>

<script>
if ("<%=rs("logo")%>"!=''){logo.innerHTML="<img border=0 src=<%=rs("logo")%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>
	<table border="0" width="97%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td height="25">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(rs("followid"))%><%=ForumTreeList%> <a href="ShowForum.asp?forumid=<%=forumid%>"><%=rs("bbsname")%></a> → 管理论坛</td>
		</tr>
	</table><br>
 
 



<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>管理论坛</b></td>
</tr>
<tr>
<td colspan="2" height="25" align="left" bgcolor="#FFFFFF"><b>&nbsp;管理帖子</b></td>
</tr>
<tr class=a3>
<td align="right" width="20%"><form name="form" method="post" action="supervise.asp?menu=delusertopic&forumid=<%=forumid%>">
清除用户 </td>
<td align="left" width="77%">
<input size="30" name="username"> 的所有主题
<input type="submit" value=" 确 定" name="Submit0"></td></form>
</tr>
<tr class=a4>
<td align="right" width="20%"><form name="form" method="post" action="supervise.asp?menu=deltopic&forumid=<%=forumid%>">
删除主题包含有 </td>
<td align="left" width="77%">
<input size="30" name="topic"> 的所有帖子
<input type="submit" value=" 确 定" name="Submit1"></td></form>
</tr>
<form name="form2" method="post" action="supervise.asp">
<input type=hidden name=menu value="up">
<input type=hidden name=forumid value="<%=forumid%>">
<tr>
<td colspan="2" height="25" align="left" bgcolor="#FFFFFF"><b>&nbsp;论坛信息</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">论坛名称<b>：</b></td>
<td class=a3 height="5" align="left" valign="middle" width="77%">
<input type="text" name="bbsname" size="30" maxlength="12" value="<%=rs("bbsname")%>">
</td>
</tr>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="20%">论坛版主<b>：</b></td>
<td class=a4 height="2" align="left" valign="middle" width="77%">
<input size="30" name="moderated" value="<%=rs("moderated")%>">
多版主添加请用|分隔，如：yuzi|裕裕
</td>
</tr>
<tr class=a3>
<td height="2" align="right" width="20%">论坛介绍<b>：</b></td>
<td height="2" align="left" valign="middle" width="77%">
<textarea name="intro" rows="4" cols="50"><%=rs("intro")%></textarea>&nbsp;
</td>
</tr>
<tr>
<td class=a4 height="1" align="right" valign="middle" width="20%">小图标URL：</td>
<td class=a4 height="1" align="left" valign="middle" width="77%">
<input size="30" name="icon" value="<%=rs("icon")%>">　显示在社区首页论坛介绍右边</td>
</tr>
<tr>
<td class=a3 align="right" valign="bottom" width="20%">大图标URL：</td>
<td class=a3 align="left" valign="bottom" width="77%">
<input size="30" name="logo" value="<%=rs("logo")%>">　显示在论坛左上角</td>
</tr>
<tr>
<td align="right" width="98%" colspan="2" bgcolor="#FFFFFF">
<input type="submit" value=" 更 新 &gt;&gt;下 一 步 " name="Submit"> </td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<Script>
document.form2.intro.value = unybbcode(document.form2.intro.value);
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>
<%
rs.close
htmlend
%>