<!-- #include file="setup.asp" -->
<%
top
id=int(Request("id"))
forumid=Conn.Execute("Select forumid From forum where id="&id)(0)
sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
bbsname=rs("bbsname")
moderated=rs("moderated")
logo=rs("logo")
followid=rs("followid")
pass=rs("pass")
password=rs("password")
userlist=rs("userlist")
Set Rs = Nothing


if Request("action")="next" then
sql="select top 1 * from forum where id > "&id&" and forumid="&forumid&" and deltopic<>1"
elseif Request("action")="Previous" then
sql="select top 1 * from forum where id < "&id&" and forumid="&forumid&" and deltopic<>1 order by id Desc"
else
sql="select * from forum where ID="&id&""
end if
Set Rs=Conn.Execute(sql)

if rs.eof or rs.bof then error"<li>系统不存在该帖子的资料"
if rs("deltopic")=1 and membercode<4 then error"<li>该主题不存在！"
topic=ReplaceText(rs("topic"),"<[^>]*>","")
%>
<!-- #include file="inc/validate.asp" -->
<script src="inc/birth.js"></script>
<script src="showad.asp"></script>
<title><%=topic%> - Powered By BBSxp</title>
<script>
function copyText(obj) {var rng = document.body.createTextRange();rng.moveToElementText(obj);rng.select();rng.execCommand('Copy');}
var i=0;
var badlist= "<%=badlist%>|<%=Request.Cookies("badlist")%>";
var moderated= "|<%=moderated%>|";
var topic="<%=topic%>"
var topicid="<%=Request("ID")%>"
if ("<%=logo%>"!=''){logo.innerHTML="<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>

<table border=0 width=97% align=center cellspacing=1 cellpadding=4 class=a2>
<tr class=a3>
<td height=25>&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?forumid=<%=forumid%>><%=bbsname%></a> → <%=topic%></td>
</tr>
</table><br>

<table cellspacing="0" cellpadding="0" width="97%" align="center" border="0"><tr><td height="35" valign="bottom"><a href="newtopic.asp?forumid=<%=rs("forumid")%>"><img alt="发表一个新主题" src="images/skins/<%=Request.Cookies("skins")%>/post.gif" border="0"></a>　<a href="retopic.asp?id=<%=rs("id")%>&topic=<%=topic%>"><img alt="回复帖子" src="images/skins/<%=Request.Cookies("skins")%>/reply.gif" border="0"></a></td><td align="right" height="35" valign="bottom"><font color="333333">您是本帖第 <b><%=rs("views")%></b> 个阅读者</font>　　<a href="?action=Previous&id=<%=rs("id")%>"><img height="12" alt="浏览上一篇主题" src="images/prethread.gif" width="52" border="0"></a>　<a style="text-decoration: none" href="javascript:location.reload()"><img height="12" alt="刷新本主题" src="images/refresh.gif" width="40" border="0"></a>　<a href="?action=next&id=<%=rs("id")%>"><img height="12" alt="浏览下一篇主题" src="images/nextthread.gif" width="52" border="0"></a></font></td></tr></table>

<SCRIPT>valigntop()</SCRIPT>

<table width="97%" border="0" cellspacing="1" class="a2" height="21" align="center"><tr class="a1">
	<td width="100%" height="18">

	<table border="0" width="100%" cellspacing="0">
		<tr class="a1">
			<td><b>&nbsp;主题</b>：<%=topic%></td>
				<td align="right"> 
<a target=_blank href=print.asp?id=<%=id%>><img alt="适合打印机打印的版本" src=images/print.gif border=0></a>&nbsp; 
<script>document.write("<a target=_blank href='mailto:?subject=<%=topic%>&body="+location.href+"'>");</script><img alt="通过电子邮件发送此页面" src=images/sendmail.gif border=0></a>&nbsp; 
<a href="javascript:window.external.AddFavorite(location.href,document.title)"><img alt="添加加到IE收藏夹" src="images/favs.gif" border="0"></a>&nbsp; 
<script>document.write("<a style=cursor:hand onclick=\"javascript:open('message.asp?menu=post&report=1&moderated=<%=moderated%>&body=【问题帖子】："+location.href+"','','width=320,height=170')\">");</script><img alt="报告本帖" src="images/feedback.gif" border="0"></a>&nbsp; 
</td></tr></table>
</td></tr></table>


<%
if Request("topage")<2 then
content=rs("content")
'''''''投票''''''''
if rs("polltopic")<>"" then
if rs("multiplicity")=1 then
multiplicity="checkbox"
else
multiplicity="radio"
end if
content=""&content&"<form action=postvote.asp?id="&rs("id")&" method=POST><table><tr><td><table>"
vote=split(rs("polltopic"),"|")
for i = 0 to ubound(vote)
if not vote(i)="" then
content=""&content&"<tr><td height=22 valign=bottom>"&i+1&".<input type="&multiplicity&" value="&i&" name=postvote id=postvote"&i&"><label for=postvote"&i&">"&vote(i)&"</label></td></tr>"
end if
next
content=""&content&"</table></td><td><table>"
allticket=0
vote=split(rs("pollresult"),"|")
for i = 0 to ubound(vote)
if not vote(i)="" then
content=""&content&"<tr><td height=22 valign=bottom>票数："&vote(i)&"</td></tr>"
allticket=vote(i)+allticket
end if
next
content=""&content&"</table></td><td><table>"
vote=split(rs("pollresult"),"|")
for i = 0 to ubound(vote)
if not vote(i)="" and allticket<>0 then
content=""&content&"<tr><td height=22 valign=bottom><img src=images/bar/0.gif width="&vote(i)/allticket*100&" height=10> ["&formatnumber(vote(i)/allticket*100)&"%]</td></tr>"
end if
next
content=""&content&"</table></td></tr><TR><TD align=center><INPUT type=submit value=投　票></TD><td colspan=5 align=center><a href=postvote.asp?menu=look&id="&id&">[查看近日参与投票的用户列表]</a></td></TR></table></form>"
end if


'''''''结束投票''''''''
sql="select * from [user] where username='"&rs("username")&"'"
Set Rs1=Conn.Execute(sql)
%>
<script>ShowPost('yuzi','<%=rs("username")%>','<%=content%>','<%=rs("posttime")%>','<%=rs1("honor")%>','<%=rs1("userface")%>','<%=rs1("sex")%>','<%=rs1("birthday")%>','<%=rs1("experience")%>','<%=rs1("membercode")%>','<%=rs1("faction")%>','<%=rs1("consort")%>','<%=rs1("money")%>','<%=rs1("posttopic")+rs1("postrevert")%>','<%=rs1("regtime")%>','<%=rs1("userlife")%>','<%=rs1("usermail")%>','<%=rs1("userhome")%>','<%=rs1("sign")%>');</script>
<%
end if
Set Rs1 = Nothing

id=rs("id")
act=topic
replies=rs("replies")
locktopic=rs("locktopic")
goodtopic=rs("goodtopic")
toptopic=rs("toptopic")
ReList=rs("ReList")

rs.close


if replies > 0 then


pagesetup=15 '设定每页的显示数量
TotalPage=cint(replies/pagesetup)  '总页数
if TotalPage < replies/pagesetup then TotalPage=TotalPage+1
PageCount = cint(Request.QueryString("ToPage"))
if PageCount < 1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage


if PageCount<2 then
sql="select top "&pagesetup&" * from [reforum"&ReList&"] where topicid="&id&" order by id"
Set Rs=Conn.Execute(sql)
else
sql="select * from [reforum"&ReList&"] where topicid="&id&" order by id"
rs.Open sql,Conn,1
end if


on error resume next  '找不到回帖时候忽略错误
if TotalPage>1 then RS.Move (PageCount-1) * pagesetup


if RS.EOF then
replies=conn.execute("Select count(id)from [reforum"&ReList&"] where topicid="&id&"")(0)
conn.execute("update [forum] set replies="&replies&" where id="&id&"")
end if



i=0
Do While Not RS.EOF and i<pagesetup
i=i+1


sql="select * from [user] where username='"&rs("username")&"'"
Set Rs1=Conn.Execute(sql)
'if RS.EOF then conn.execute("delete from [reforum"&ReList&"] where username='"&rs("username")&"'")
%><script>var i=<%=i+(PageCount-1)*pagesetup%>;ShowPost('<%=rs("id")%>','<%=rs("username")%>','<%=rs("content")%>','<%=rs("posttime")%>','<%=rs1("honor")%>','<%=rs1("userface")%>','<%=rs1("sex")%>','<%=rs1("birthday")%>','<%=rs1("experience")%>','<%=rs1("membercode")%>','<%=rs1("faction")%>','<%=rs1("consort")%>','<%=rs1("money")%>','<%=rs1("posttopic")+rs1("postrevert")%>','<%=rs1("regtime")%>','<%=rs1("userlife")%>','<%=rs1("usermail")%>','<%=rs1("userhome")%>','<%=rs1("sign")%>');</script>
<%
Set Rs1 = Nothing


RS.MoveNext
loop
RS.Close


end if


if TotalPage=0 then
TotalPage=1
PageCount=1
end if
%>
<SCRIPT>


//function window.onload(){if(eval("div"+id+".scrollHeight")>100){eval("div"+id+".style.height=100")}}

valignbottom()</SCRIPT>

<table cellspacing="4" cellpadding="0" width="97%" border="0" style="border-collapse: collapse" align=center><tr><td width="61%"><b>本主题共有 <font color="990000"><%=TotalPage%></font> 页 
	[
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"id=<%=id%>")
</script>

 ]</b></td><td width="39%" align="right">
<a href="favorites.asp?menu=add&url=topic&name=<%=id%>">收藏帖子</a> | <a href="favorites.asp?menu=del&url=topic&name=<%=id%>">取消收藏</a> | <a href="#">返回页首</a>&nbsp;</td></tr></table>
<%if Request.Cookies("username")<>"" then%>

<form name="yuziform" method="post" action="retopic.asp" onSubmit="return CheckForm(this);">
<input name="content" type="hidden">
<input type="hidden" value="<%=id%>" name="id">
<input type="hidden" value="<%=act%>" name="topic">
<SCRIPT>valigntop()</SCRIPT>

<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align=center>
<tr><td width="20%" height="20" class="a1"><b>&nbsp;快速回复主题</b></td><td class="a1" width="80%"><b><%=act%></b></td></tr>
<tr>
	<td width="20%" class="a3" rowspan="2">

<b>&nbsp;用户名</b>：<%=Request.Cookies("username")%> [<a href="login.asp?menu=out"><font color="000000">退出</font></a>]
<hr>
<TABLE cellSpacing=1 cellPadding=3 border=0>
<TR>
<script>
ii=0
for(i=1;i<=25;i++) {
index = Math.floor(Math.random() * 80+1);
ii=ii+1
document.write("<TD><a href=javascript:emoticon('"+index+"')><img border=0 src=images/Emotions/"+index+".gif></a></TD>")
if (ii==5){document.write("</TR><TR align=middle>");ii=0;}
}
</script>
</TR></TABLE>

	<td width="80%" height="150" class="a4">


<SCRIPT src="inc/post.js"></SCRIPT>
 <br>　</td></tr><tr>
			<td valign="top" class="a4"><input tabindex="4" type="submit" value="Ctrl+Enter 回复主题" name="submit1">　　<input name=preview type="Button" value=" 预 览 " onclick="Gopreview()">　　<input type="reset" name="reset" value=" 重 置 "></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<form name=preview action=preview.asp method=post target=preview_page><input name="content" type="hidden"></form>
<table cellspacing="0" cellpadding="0" width="97%" align="center" border="0"><tr><td align="middle">管理选项: <%
response.write "<a href=manage.asp?menu=movenew&id="&id&">拉前主题</a> | "
if locktopic=1 then
response.write "<a href=manage.asp?menu=dellocktopic&id="&id&">开放主题</a>"
else
response.write "<a href=manage.asp?menu=locktopic&id="&id&">关闭主题</a>"
end if
response.write " | "
if toptopic=2 then
response.write "<a href=manage.asp?menu=untop&id="&id&">取消总置顶</a>"
else
response.write "<a href=manage.asp?menu=top&id="&id&">主题总置顶</a>"
end if
response.write " | "
if toptopic=1 then
response.write "<a href=manage.asp?menu=deltoptopic&id="&id&">取消置顶</a>"
else
response.write "<a href=manage.asp?menu=toptopic&id="&id&">主题置顶</a>"
end if
response.write " | "
if goodtopic=1 then
response.write "<a href=manage.asp?menu=delgoodtopic&id="&id&">移出精华区</a>"
else
response.write "<a href=manage.asp?menu=goodtopic&id="&id&">添加到精华区</a>"
end if
%>| <a href="move.asp?id=<%=id%>">移动主题</a> | <a href="manage.asp?menu=deltopic&id=<%=id%>">删除主题</a> </td></tr></table><%end if%> 

<!-- #include file="inc/line.asp" -->
<%
conn.execute("update [forum] set Views=Views+1 where id="&id&"")

htmlend
%>