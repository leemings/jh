<!-- #include file="setup.asp" -->
<%
top

forumid=int(Request("forumid"))
sql="select * from bbsconfig where id="&forumid&""
Set Rs=Conn.Execute(sql)
bbsname=rs("bbsname")
toltopic=rs("toltopic")
moderated=rs("moderated")
logo=rs("logo")
followid=rs("followid")
hide=rs("hide")
pass=rs("pass")
password=rs("password")
userlist=rs("userlist")

rs.close

%>


<!-- #include file="inc/validate.asp" -->

<meta http-equiv=refresh content=300>
<script>
var key_word="<%=Request.Cookies("key_word")%>"
var cookieusername="<%=Request.Cookies("username")%>"
if ("<%=logo%>"!=''){logo.innerHTML="<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'>"}
</script>


<title><%=bbsname%> - Powered By BBSxp</title>
	<table border="0" width="97%" align="center" cellspacing="1" cellpadding="4" class=a2>
		<tr class=a3>
			<td colspan=2><table border="0" width="100%" cellspacing="0" cellpadding="0"><tr>
				<td height="18">&nbsp;<img src=images/Forum_nav.gif>&nbsp; <%ClubTree%> → <%ForumTree(followid)%><%=ForumTreeList%> <a href=ShowForum.asp?forumid=<%=forumid%>><%=bbsname%></a></td>
				<td height="18" align="right"><img src=images/jt.gif> <a href=supervise.asp?forumid=<%=forumid%>>论坛管理</a>
</td></table>
</td>
</tr>
<%
ShowForum=1
sql="Select * From bbsconfig where followid="&forumid&" and hide=0 order by SortNum"
Set Rs1=Conn.Execute(sql)
if not Rs1.eof then
do while not rs1.eof
if ""&Rs1("password")&""="" then:password=0:else:password=1:end if
%><script>ShowList("<%=Rs1("pass")%>","<%=Rs1("id")%>","<%=Rs1("bbsname")%>","<%=Rs1("icon")%>","<%=Rs1("intro")%>","<%=Rs1("moderated")%>","<%=Rs1("today")%>","<%=Rs1("toltopic")%>","<%=Rs1("tolrestore")%>","<%=Rs1("lasttime")%>","<%=Rs1("lastname")%>","<%=Rs1("lasttopic")%>","<%=password%>");</script>
<%
rs1.movenext
loop
Set Rs1 = Nothing

ShowForum=0
end if


%>
</table>

<br>

<!-- #include file="inc/line.asp" -->
<%

if ShowForum=1 or SortShowForum=1 then

forumidonline=conn.execute("Select count(sessionid)from online where forumid="&forumid&"")(0)
regforumidonline=conn.execute("Select count(sessionid)from online where forumid="&forumid&" and username<>''")(0)
%><table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr><td width="93%" height="25" class="a1">　<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=forumid%>)"> 目前论坛总在线  <b><%=onlinemany%></b> 人，本分论坛共有 <b><%=forumidonline%></b> 人在线。其中注册用户 <b><%=regforumidonline%></b> 人，访客 <b><%=forumidonline-regforumidonline%></b> 人。</td><td align="middle" width="7%" height="25" class="a1"><a href="javascript:this.location.reload()"><img src="images/refresh.gif" border="0"></a></td></tr><tr height="25" style="display:none" id="follow0"><td id="followTd0" align="left" class="a4" width="94%" colspan="5">
　Loading...</td></tr></tr></table>
<br>



<table height="30" cellspacing="0" cellpadding="0" width="97%" align="center" border="0"><tr><td align="left" width="20%"><a href="newtopic.asp?forumid=<%=forumid%>"><img src="images/skins/<%=Request.Cookies("skins")%>/post.gif" border="0" alt="发表文章"></a></td><td align="right" width="80%">
<img src=images/showdigest.gif> <a onmouseover="showmenu(event,'<div class=menuitems><a href=favorites.asp?menu=add&url=forum&name=<%=forumid%>>收藏论坛</a></div><div class=menuitems><a href=favorites.asp?menu=del&url=forum&name=<%=forumid%>>取消收藏</a></div>')" style=cursor:default>论坛收藏</a>
<SCRIPT>
var moderated="<%=moderated%>"
var listmoderated=""
var list= moderated.split ('|'); 
for(i=0;i<list.length;i++) {
if (list[i] !=""){
listmoderated=listmoderated+"<div class=menuitems><a href=profile.asp?username="+list[i]+">"+list[i]+"</a></div>";
}
}
document.write("<img src=images/team.gif> <a onmouseover=\"showmenu(event,'"+listmoderated+"')\" style=cursor:default>论坛版主</a>");
</SCRIPT>
</td></tr></table><table height="28" cellspacing="1" cellpadding="1" width="97%" align="center" border="0"><tr><td align="middle" width="3%"><img height="18" src="images/announce.gif" width="18" align="middle" alt="社区公告"></td>
		<td><marquee onmouseover="this.stop()" onmouseout="this.start()" width="90%" scrollamount="3"><a target=_blank href="affiche.asp"><%=affichetitle%></a></marquee></td>
		<td align="right" width="200">
	&nbsp;<img src=images/jt.gif> <a onmouseover="showmenu(event,'<div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=1>查看一天内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=3>查看三天内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=7>查看一星期内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=14>查看两星期内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=30>查看一个月内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=90>查看三个月内的主题</a></div><div class=menuitems><a href=?forumid=<%=forumid%>&TimeLimit=180>查看六个月内的主题</a></div>')" href=?forumid=<%=forumid%>>
	查看所有主题</a>
	<img src=images/jt.gif> <a href="ShowForum.asp?forumid=<%=forumid%>&search=goodtopic">
查看精华帖子</a>
			


</td></tr></table>


<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr height="25" id=TableTitleLink><td width="3%" class="a1"><td width="3%" class="a1"><td align="middle" height="24" class="a1" width="45%"><a href="ShowForum.asp?forumid=<%=forumid%>&order=id&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">主题</a></td><td align="middle" width="9%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=username&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">作者</a></font></td><td align="middle" width="6%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=replies&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">回复</a></td><td align="middle" width="7%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=Views&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">点击</a></td><td width="20%" height="24" class="a1" align="center"><a href="ShowForum.asp?forumid=<%=forumid%>&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">最后更新</a></td></tr>


<%

if Request("order")<>"" then
order=HTMLEncode(Request("order"))
else
order="lasttime"
end if


if Request("TimeLimit")<>"" then TimeLimit="and lasttime>"&SqlNowString&"-"&int(Request("TimeLimit"))&""

if Request("search")="goodtopic" then search="and goodtopic=1 "

if Request.Cookies("pagesetup")=empty then
pagesetup=20   '设定每页的显示数量
else
pagesetup=int(Request.Cookies("pagesetup"))
if pagesetup > 30 then pagesetup=20
end if

topsql="where deltopic<>1 and forumid="&forumid&" "&search&" "&TimeLimit&" or toptopic=2"



if Request("TimeLimit")<>"" or Request("search")<>"" then
count=conn.execute("Select count(id) from [forum] "&topsql&"")(0)
else
count=toltopic
end if


TotalPage=0
TotalPage=cint(count/pagesetup)  '总页数
if TotalPage < count/pagesetup then TotalPage=TotalPage+1
PageCount = cint(Request.QueryString("ToPage"))
if PageCount < 1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage

if PageCount<2 then
sql="select top "&pagesetup&" * from [forum] "&topsql&" order by toptopic Desc,"&order&" Desc"
Set Rs=Conn.Execute(sql)
else
sql="select * from [forum] "&topsql&" order by toptopic Desc,"&order&" Desc"
rs.Open sql,Conn,1
end if

if TotalPage>1 And Not RS.EOF then RS.Move (PageCount-1) * pagesetup

if RS.EOF and toltopic>0 then UpTolTopic



i=0
Do While Not RS.EOF and i<pagesetup
i=i+1

newtopic=""
if rs("posttime")+1>now() then newtopic="<img src=images/new.gif>"

%><script>ShowForum("<%=rs("ID")%>","<%=rs("topic")%>","<%=newtopic%>","<%=rs("username")%>","<%=rs("Views")%>","<%=rs("icon")%>","<%=rs("toptopic")%>","<%=rs("locktopic")%>","<%=rs("pollresult")%>","<%=rs("goodtopic")%>","<%=rs("replies")%>","<%=rs("lastname")%>","<%=rs("lasttime")%>");</script>
<%
RS.MoveNext
loop
RS.Close
%></table>
<SCRIPT>valignbottom()</SCRIPT>
<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0"><tr height="25"><td width="100%" height="2"><table cellspacing="0" cellpadding="3" width="100%" border="0"><tr><td height="2">
		<b>本论坛共有 <font color=990000><%=TotalPage%></font> 页 [ <b>
<script>
ShowPage(<%=TotalPage%>,<%=PageCount%>,"forumid=<%=forumid%>&order=<%=Request("order")%>&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>")
</script>

</b>]

</b></td><form name="form" action="search.asp?menu=ok&forumid=<%=forumid%>&search=key&searchxm2=topic" method="post"><input type=hidden name=sessionid value=<%=session.sessionid%>>
<td height="2" align="right">快速搜索：<input name="content" size="20" onkeyup="ValidateTextboxAdd(this, 'btnadd')" onpropertychange="ValidateTextboxAdd(this, 'btnadd')" onfocus="javascript:focusEdit(this)" onblur="javascript:blurEdit(this)" value="关键词" helptext="关键词"> <input type="submit" value="搜索" name="submit" id='btnadd' disabled> </td></form></tr></table></td></tr></table>

　<table cellspacing="4" cellpadding="0" width="80%" border="0" align=center><tr><td nowrap width="200"><img alt="" src="images/f_new.gif" border="0"> 打开主题 (有回复的主题)</td><td nowrap width="100"><img alt="" src="images/f_hot.gif" border="0"> 热门主题</td><td nowrap width="100"><img alt="" src="images/f_locked.gif" border="0"> 关闭主题</td><td nowrap width="150"><img src="images/topicgood.gif"> 精华主题</td></tr><tr><td nowrap width="200"><img alt="" src="images/f_norm.gif" border="0"> 打开主题 (没有回复的主题)</td><td nowrap width="100"><img alt="" src="images/f_poll.gif" border="0"> 投票主题</td><td nowrap width="100"><img alt="" src="images/f_top.gif" border="0"> 置顶主题</td><td nowrap width="150"><img src="images/my.gif"> 自己发表的主题</td></tr></table></center></div>
<%

end if


htmlend


sub UpTolTopic
allarticle=conn.execute("Select count(forumid) from forum where deltopic<>1 and forumid="&forumid&"")(0)
if allarticle>0 then
allrearticle=conn.execute("Select sum(replies) from forum where deltopic<>1 and forumid="&forumid&"")(0)
else
allrearticle=0
end if
conn.execute("update [bbsconfig] set toltopic="&allarticle&",tolrestore="&allarticle+allrearticle&" where ID="&forumid&"")
end sub

%>