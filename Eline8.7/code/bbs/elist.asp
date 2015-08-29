<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<%
'=========================================================
' File: elist.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

	dim Ers,Esql
	dim totalrec
	dim currentpage,page_count,Pcount
	dim guests
	dim topic
	dim rs1,sql1
	dim endpage
	stats="精华帖子列表"
	currentPage=request("page")
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if currentpage="" then
		currentpage=1
	elseif not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if

	if founderr then
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	else
		call nav()
		call head_var(1,BoardDepth,0,0)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()

sub main()
	if Cint(GroupSetting(41))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛精华帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
			end if
		end if
	end if
	'###################特殊版面修改开始(asilas制作)##################
	if cint(Board_Setting(51))=1 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkviplogin(membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
			founderr=true
			exit sub
			end if
		end if
	end if
	if cint(Board_Setting(52))<>0 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			dim sexshow
			if cint(Board_Setting(52))=1 then
			sexshow="女生"
			elseif cint(Board_Setting(52))=2 then
			sexshow="男生"
			end if
			if chksexlogin(cint(Board_Setting(52)),membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>"&sexshow&"论坛版面</font>，请确认您的性别是否符合。"
			founderr=true
			exit sub
			end if
		end if
	end if
	'####################特殊版面修改结束(asilas制作)#################
	call activeonline()
	dim onlinenum,guestnum
	onlinenum=online(boardid,false)
	guestnum=guest(boardid,false)
	if trim(boardmasterlist)<>"" then
		master_1=split(boardmasterlist, "|")
		for i = 0 to ubound(master_1)
		master_2=""+master_2+"<a href=""dispuser.asp?name="+master_1(i)+""" target=_blank title=点击查看该版主资料>"+master_1(i)+"</a>&nbsp;"
		next
	else
		master_2="无版主"
	end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=100% valign=middle><!--#include file="z_announce.asp"--></td>
</tr>
</table>
<BR>
<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TBODY>
<TR>
<Th height=25 width="100%" align=left id=tabletitlelink style="font-weight:normal">总在线<b><%=allonline()%></b>人，其中<%= boardtype %>上共有 <b><%= onlinenum %></b> 位会员与 <b><%= guestnum %></b> 位客人.今日帖子 <b><font color=<%=Forum_body(8)%>><%= todaynum %></font></b> 
<%
	if request("action")="show" then
		response.write "[<a href=elist.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
	else
		if cint(Forum_Setting(14))=1 and request("action")<>"off" then
		response.write "[<a href=elist.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
		else
		response.write "[<a href=elist.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">显示详细列表</a>]"
		end if
	end if
%>
</Th></TR>
<%
	if request("action")="off" then
		call onlineuser(0,0,boardid)
	elseif request("action")="show" then
		call onlineuser(1,1,boardid)
	else
		call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
	end if
%>
</td></tr></TBODY></TABLE>
<BR>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center valign=middle><tr>
<td align=center width=2> </td>
<td align=left> <a href="announce.asp?boardid=<%= boardid %>"><img src="<%=Forum_info(7)&Forum_boardpic(1)%>" border=0 alt="发新帖"></a>
&nbsp;&nbsp;<a href="vote.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(2)%>" border=0 alt="发起新投票"></a>
&nbsp;&nbsp;<a href="smallpaper.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(3)%>" border=0 alt="发布小字报"></a></td>
<td align=right><img src="<%=Forum_info(7)%>team2.gif" align=absmiddle>  <%= master_2 %></td></tr></table>
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TR valign=middle>
<Th height=27 width=32>状态</Th>
<Th width=*>主 题</Th>
<Th width=80>作 者</Th>
<Th width=70>最后更新</Th>
<Th width=80>回复人</Th>
</TR>
<%
	sql="select b.rootid,t.* from BestTopic b inner join topic t on b.rootid=t.topicid where t.boardID="&cstr(boardID)&" ORDER BY b.id desc"
	'response.write sql
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
	response.write "<tr><td colSpan=5 width=100% class=tablebody1>本论坛暂无精华帖子</td></tr>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
'========================Best Topic Info=======================
response.write "<TR valign=middle>"
response.write "<TD class=tablebody2 width=32 height=27 align=center>"
response.write TopicIcon(rs("BoardID"),rs("topicid"),rs("lastposttime"),rs("istop"),rs("isVote"),rs("LockTopic"),iif(Rs("child")>=Cint(forum_setting(44)),1,0))
response.write "</TD>"
response.write "<TD align=left onmouseover=this.className='TableBody2' onmouseout=this.className='TableBody1' class=tablebody1 width=*>"
response.write "&nbsp;<img src='"& Forum_info(8) & rs("Expression") &"'>&nbsp;"
response.write "<a href=""dispbbs.asp?boardID="& boardID &"&ID="& rs("rootid") &"&skin=1"">"
if len(rs("title"))>26 then
	response.write htmlencode(left(reubbcode(rs("title"),false),26))&"..."
else
	response.write htmlencode(reubbcode(rs("title"),false))
end if
response.write "</a>"
response.write "&nbsp;&nbsp;<img src="""&Forum_info(7)&Forum_statePic(4)&""">"
response.write "</TD>"
response.write "<TD class=tablebody2 width=80 align=center><a href=dispuser.asp?id="& rs("postuserid") &" target=_blank>"& rs("postusername") &"</a></TD>"
response.write "<TD align=center class=tablebody1 width=70>"&month(rs("dateandtime"))&"-"&day(rs("dateandtime"))&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&"</TD>"
response.write "<TD align=center class=tablebody2 width=80>"
if conn.execute("select count(*) from topic where topicid=" & rs("rootid"))(0) > 0 then
 response.write split(conn.execute("select lastpost from topic where topicid=" & rs("rootid"))(0),"$")(0)
else
 response.write "------"
end if
response.write "</TD>"
response.write "</TR>"
'============================End=============================
		page_count = page_count + 1
		rs.movenext
		wend
	end if
  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
%>
</table>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr><td valign=middle nowrap>
页次：<b><%=currentpage%></b>/<b><%=Pcount%></b>页
&nbsp;
每页<b><%=Forum_Setting(11)%></b> 总数<b><%=totalrec%></b></td>
<td valign=middle nowrap align=right>分页： 
<%call DispPageNum(currentpage,PCount,"""?page=","&boardid="&boardid&"""")
	response.write "</td></tr></table>"
	rs.close
	set rs=nothing
%>
<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr>
<td valign=middle nowrap> <div align=right>
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option selected>跳转论坛至...</option>
<%
myCache.name="BoardJumpList"
if myCache.valid then
response.write myCache.value
else
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
if boardid=rs(0) then
BoardJumpList = BoardJumpList & "selected"
end if
BoardJumpList = BoardJumpList & ">"
select case rs(2)
case 0
BoardJumpList = BoardJumpList & "╋"
case 1
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end select
if rs(2)>1 then
for i=2 to rs(2)
	BoardJumpList = BoardJumpList & "&nbsp;&nbsp;│"
next
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end if
BoardJumpList = BoardJumpList & rs(1)&"</option>"
rs.MoveNext
loop
set rs=nothing
myCache.add BoardJumpList,dateadd("n",60,now)
response.write BoardJumpList
end if
%>
</select><div></td></tr></table>
<BR>
<%response.write "<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>"&_
	"<th width=80% align=left>　-=> "& Forum_info(0) &"图例</th>"&_
	"<th noWrap width=20% align=right>所有时间均为 - "&Forum_info(9)&" &nbsp;</th>"&_
	"</tr>"&_
	"<tr><td colspan=2 class=tablebody1>"&_
	"<table cellspacing=4 cellpadding=0 width=90% border=0 align=center>"&_
	"<tr>"&_
	"<td><img src=pic/topicicon/00000.gif> 旧主题</td>"&_
	"<td><img src=pic/topicicon/10000.gif> 新主题</td>"&_
	"<td><img src=pic/topicicon/00001.gif> 热门主题</td>"&_
	"<td><img src=pic/topicicon/00010.gif> 锁定主题</td>"&_
	"<td><img src=pic/topicicon/00100.gif> 投票</td>"&_
	"<td><img src=pic/topicicon/01000.gif> 固顶主题</td>"&_
	"<td><img src=pic/topicicon/02000.gif> 总固顶主题</td>"&_
	"</tr>"&_
	"</table>"&_
	"</td></tr></table>"
end sub
%>
<!--#include file="online_l.asp"-->