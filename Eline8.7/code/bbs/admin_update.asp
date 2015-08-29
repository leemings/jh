<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
Server.ScriptTimeout=9999999
dim admin_flag
admin_flag="14,28"
if not master or instr(session("flag"),"14")=0 or instr(session("flag"),"28")=0 then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	dim tmprs
	dim body
	call main()
	conn.close
	set conn=nothing
end if

sub main()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=2 height=23>论坛数据处理</th>
</tr>
<tr>
<td width="20%" class="forumrow">注意事项</td>
<td width="80%" class="forumrow">下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。</td>
</tr>
<%
	if request("action")="updat" then
		if request("submit")="更新论坛数据" then
		call updateboard()
		elseif request("submit")="修 复" then
		call fixtopic()
		else
		call updateall()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	elseif request("action")="delboard" then
		if isnumeric(request("boardid")) then
		conn.execute("update topic set locktopic=2 where boardid="&request("boardid"))
		for i=0 to ubound(AllPostTable)
		conn.execute("update "&AllPostTable(i)&" set locktopic=2 where boardid="&request("boardid"))
		next
		end if
		response.write "<tr><td align=left colspan=2 height=23 class=forumrow>清空论坛数据成功，请返回更新帖子数据！</td></tr>"
	elseif request("action")="updateuser" then
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户发帖</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>发帖重新计算所有用户发表帖子数量。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="重新计算用户发帖"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户金钱/经验/魅力</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户的发帖数量和论坛的相关设置重新计算用户的金钱/经验/魅力，本操作也将重新计算贵宾、版主、总版主的数据<BR>注意：不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发帖和总的论坛分值设置进行运算，请大家慎重操作，<font color=red>而且本项操作将重置用户因为奖励、惩罚等原因管理员对用户分值的修改。</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户金钱/经验/魅力"></td>
</tr>
</FORM>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户威望</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖时得到的鲜花和鸡蛋数量重新计算用户的威望。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户威望"></td>
</tr>
</FORM>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户积分</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖时得到的积分数量重新计算用户的积分。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户积分"></td>
</tr>
</FORM>
<%
	elseif request("action")="updateuserinfo" then
		if request("submit")="重新计算用户发帖" then
			call updateTopic()
		elseif request("submit")="更新用户等级" then
			call updategrade()
		elseif request("submit")="更新用户威望" then
			call updatepower()
		elseif request("submit")="更新用户积分" then
			call updatescore()
		else
			call updatemoney()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	else
%>
<tr> 
<th align=left colspan=2 height=23>更新论坛数据</th>
</tr>

<form action="admin_update.asp?action=updat" method=post>
<tr>
<td width="20%" class="forumrow">更新分论坛数据</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新论坛数据"><BR><BR>这里将重新计算每个论坛的帖子主题和回复数，今日帖子，最后回复信息等，建议每隔一段时间运行一次。</td>
</tr>
<tr>
<td width="20%" class="forumrow">更新总论坛数据</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新论坛总数据"><BR><BR>这里将重新计算整个论坛的帖子主题和回复数，今日帖子，最后加入用户等，建议每隔一段时间运行一次。</td>
</tr>
<tr> 
<th align=left colspan=2 height=23>修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="1000" size=10>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 复"></td>
</tr>
</form>
<%
	end if
%>
</table><BR><BR>
<%
	end sub

sub updateboard()
	'先按照所有版面ID得出帖子数，然后计算各个有下属论坛的帖子总和
	dim allarticle
	dim alltopic
	dim alltoday
	dim allboard
	dim trs,Esql,ars
	dim Maxid
	dim LastTopic,LastRootid,LastPostTime,LastPostUser
	dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	dim ParentStr
	set rs=server.createobject("adodb.recordset")
	if isnumeric(request("boardid")) and request("boardid")<>"" then
		sql="select boardid,boardtype from board where boardid="&request("boardid")
	else
		sql="select boardid,boardtype from board"
	end if
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<tr><td align=left colspan=2 height=23 class=forumrow>还没有论坛版面，请到论坛管理添加！</td></tr>"
		founderr=true
		exit sub
	else
		do while not rs.eof
	    allarticle=BoardAnnounceNum(rs("BoardID"))
	    alltopic=BoardTopicNum(rs("BoardID"))
			sql="select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.boardid="&rs("boardid")&" and b.locktopic<2 order by b.Announceid desc"
			set trs=conn.execute(sql)
			if not(trs.eof and trs.bof) then
				Lasttopic=replace(left(trs(0),15),"$","")
				LastRootid=trs(1)
				LastPostTime=trs(2)
				LastPostUser=trs(3)
				LastPostUserid=trs(4)
				Lastid=trs(5)
			else
				LastTopic="无"
				LastRootid=0
				LastPostTime=now()
				LastPostUser="无"
				LastPostUserid=0
				Lastid=0
			end if
			trs.close
			set trs=nothing
			LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & rs("BoardID")
			Esql="update board set lastbbsnum="&allarticle&",lasttopicnum="&alltopic&",TodayNum="&todays(rs("boardid"))&",LastPost='"&replace(LastPost,"'","")&"' where boardid="&rs("boardid")
			conn.execute(Esql)
			rs.movenext
		loop
		'更新有下属论坛的版面数据
		dim LastPostinfo,LastAnnID
		LastAnnID=0
		if isnumeric(request("boardid")) and request("boardid")<>"" then
			set trs=conn.execute("select boardid,ParentStr from board where boardid="&request("boardid")&" and child>0")
		else
			set trs=conn.execute("select boardid,ParentStr from board where child>0")
		end if
		do while not trs.eof
			if trs(1)="0" then
				ParentStr=trs(0)
			else
				ParentStr=trs(1) & "," & trs(0)
			end if
			set ars=conn.execute("select sum(lastbbsnum),sum(lasttopicnum),sum(todaynum),count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
			'帖子总数，主题总数，今日帖总数，下属版面数
			allarticle=ars(0)
			alltopic=ars(1)
			alltoday=ars(2)
			allboard=ars(3)
			if isnull(allarticle) then allarticle=0
			if isnull(alltopic) then alltopic=0
			if isnull(alltoday) then alltoday=0
			if isnull(allboard) then allboard=0
			set ars=conn.execute("select LastPost from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
			do while not ars.eof
				LastPostinfo=split(ars(0),"$")
				if isnumeric(LastPostInfo(1)) and isnumeric(LastAnnID) then
					if clng(LastPostInfo(1))>clng(LastAnnID) then
					LastPost=ars(0)
					else
					LastPost="$0$"&Now()&"$$$$$"
					end if
				else
					LastPost="$0$"&Now()&"$$$$$"
				end if
				LastAnnID=LastPostinfo(1)
				ars.movenext
			loop
			conn.execute("update board set lastbbsnum=lastbbsnum+"&allarticle&",lasttopicnum=lasttopicnum+"&alltopic&",TodayNum=todaynum+"&alltoday&",child="&allboard&",LastPost='"&checkStr(LastPost)&"' where boardid="&trs(0))
			'response.write AllBoard & "<br>"
			trs.movenext
		loop
	end if
	if isnumeric(request("boardid")) and request("boardid")<>"" then
		sql="select boardtype,lastbbsnum,lasttopicnum,todaynum from board where boardid="&request("boardid")
	else
		sql="select boardtype,lastbbsnum,lasttopicnum,todaynum from board"
	end if
	set rs=conn.execute(sql)
	do while not rs.eof
		body=body & "<tr><td colspan=2 class=forumrow>更新论坛数据成功，"&rs("boardtype")&"共有"&rs("lastbbsnum")&"篇帖子，"&rs("lasttopicnum")&"篇主题，今日有"&rs("todaynum")&"篇帖子。</td></tr>"
		rs.movenext
	loop
	set rs=nothing
	set trs=nothing
	set ars=nothing
end sub

sub updateall()
	sql="update config set TopicNum="&topicnum()&",BbsNum="&announcenum()&",TodayNum="&alltodays()&",UserNum="&allusers()&",lastUser='"&newuser()&"'"
	conn.execute(sql)
	body="<tr><td colspan=2 class=forumrow>更新总论坛数据成功，全部论坛共有"&announcenum()&"篇帖子，"&topicnum()&"篇主题，今日有"&alltodays()&"篇帖子，有"&allusers()&"用户，最新加入为"&newuser()&"。</td></tr>"
end sub

sub fixtopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim TotalUseTable,Ers
dim username,dateandtime,rootid,announceid,postuserid,lastpost,topic,childnum
set rs=server.createobject("adodb.recordset")
sql="select topicid,PostTable,boardid from topic where topicid>="&request.form("beginid")&" and topicid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	childnum=conn.execute("select count(*) from "&rs(1)&" where rootid="&rs(0)&" and locktopic<2")(0)
	if isnull(childnum) then childnum=0
	conn.execute("update topic set child="&(childnum-1)&" where topicid="&rs(0))
	conn.execute("update "&rs(1)&" set boardid="&rs("boardid")&" where rootid="&rs(0)&" and boardid<>"&rs("boardid"))
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body,boardid from "&rs(1)&" where rootid="&rs(0)&" and locktopic<2 order by Announceid desc"
	set ers=conn.execute(sql)
	if not (ers.eof and ers.bof) then
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		topic=left(Ers("body"),20)
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
		LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid & "$" & ers("BoardID")
		conn.execute("update topic set LastPost='"&replace(LastPost,"'","")&"' where topicid="&rs(0))
	end if
rs.movenext
loop
set ers=nothing
set rs=nothing
%>
<form action="admin_update.asp?action=updat" method=post>
<tr> 
<th align=left colspan=2 height=23>继续修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 复"></td>
</tr>
</form>
<%
end sub

'分论坛今日帖子
function todays(boardid)
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where boardid="&boardid&" and  locktopic<2 and datediff('d',dateandtime,Now())=0")
todays=tmprs(0)
set tmprs=nothing
if isnull(todays) then todays=0
end function

'全部论坛今日帖子
function alltodays()
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where locktopic<2 and datediff('d',dateandtime,Now())=0")
alltodays=tmprs(0)
set tmprs=nothing
if isnull(alltodays) then alltodays=0
end function

'所有注册用户数量
function allusers() 
tmprs=conn.execute("Select count(userid) from [user]") 
allusers=tmprs(0) 
set tmprs=nothing 
if isnull(allusers) then allusers=0 
end function
'最新注册用户
function newuser()
sql="Select top 1 username from [user] order by userid desc"
set tmprs=conn.execute(sql)
if tmprs.eof and tmprs.bof then
	newuser="没有会员"
else
   	newuser=tmprs("username")
end if
set tmprs=nothing
end function 

'所有论坛帖子
function AnnounceNum()
dim AnnNum
AnnNum=0
AnnounceNum=0
For i=0 to ubound(AllPostTable)
	tmprs=conn.execute("Select Count(announceID) from "&AllPostTable(i)&" where locktopic<2") 
	AnnNum=tmprs(0)
	set tmprs=nothing 
	if isnull(AnnNum) then AnnNum=0
	AnnounceNum=AnnounceNum + AnnNum
next
set tmprs=nothing
end function
'分论坛帖子
function BoardAnnounceNum(boardid)
dim BoardAnnNum
BoardAnnNum=0
BoardAnnounceNum=0
For i=0 to ubound(AllPostTable)
	tmprs=conn.execute("Select Count(announceID) from "&AllPostTable(i)&" where boardid="&boardid&" and locktopic<2") 
	BoardAnnNum=tmprs(0) 
	set tmprs=nothing 
	if isnull(BoardAnnNum) then BoardAnnNum=0
	BoardAnnounceNum=BoardAnnounceNum + BoardAnnNum
next
set tmprs=nothing
end function

'所有论坛主题
function TopicNum() 
tmprs=conn.execute("Select Count(topicid) from topic where locktopic<2") 
TopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(TopicNum) then TopicNum=0 
end function 
'分论坛主题
function BoardTopicNum(boardid) 
tmprs=conn.execute("Select Count(topicid) from topic where boardid="&boardid&" and locktopic<2")
BoardTopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(BoardTopicNum) then BoardTopicNum=0 
end function

'更新用户发帖数
sub updateTopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim userTopic
sql="select userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	userTopic=Userallnum(rs(0))
	conn.execute("update [user] set article="&userTopic&" where userid="&rs(0))
rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户发帖</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>发帖重新计算所有用户发表帖子数量。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="重新计算用户发帖"></td>
</tr>
</form>
<%
end sub

'更新用户威望
sub updatePower()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim userPower
sql="select userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	userPower=UserPowerNum(rs(0))
	conn.execute("update [user] set userpower="&userPower&" where userid="&rs(0))
	rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户威望</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖时得到的鲜花和鸡蛋数量重新计算用户的威望。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户威望"></td>
</tr>
</form>
<%
end sub

'更新用户积分
sub updateScore()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim userScore
sql="select userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	userScore=UserScoreNum(rs(0))
	conn.execute("update [user] set userscore="&userScore&" where userid="&rs(0))
	rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户积分</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖时得到的积分数量重新计算用户的积分。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户积分"></td>
</tr>
</form>
<%
end sub

'更新用户金钱/经验/魅力
sub updatemoney()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim userTopic,userReply,userWealth
dim userEP,userCP

sql="select logins,userid from [user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=conn.execute(sql)
do while not rs.eof
	userTopic=UserTopicNum(rs(1))
	userreply=UserReplyNum(rs(1))
	userwealth=rs(0)*Forum_user(4) + userTopic*Forum_user(1) + userreply*Forum_user(2)
	userEP=rs(0)*Forum_user(9) + userTopic*Forum_user(6) + userreply*Forum_user(7)
	userCP=rs(0)*Forum_user(14) + userTopic*Forum_user(11) + userreply*Forum_user(12)
	if isnull(UserWealth) or not isnumeric(userwealth) then userwealth=0
	if isnull(Userep) or not isnumeric(userep) then userep=0
	if isnull(Usercp) or not isnumeric(usercp) then usercp=0
	conn.execute("update [user] set userWealth="&userWealth&",userep="&userep&",usercp="&usercp&" where userid="&rs(1))
rs.movenext
loop
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>更新用户金钱/经验/魅力</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户的发帖数量和论坛的相关设置重新计算用户的金钱/经验/魅力，本操作也将重新计算贵宾、版主、总版主的数据<BR>注意：不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发帖和总的论坛分值设置进行运算，请大家慎重操作，<font color=red>而且本项操作将重置用户因为奖励、惩罚等原因管理员对用户分值的修改。</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户金钱/经验/魅力"></td>
</tr>
</form>
<%
end sub

'更新用户等级
sub updategrade()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim oldMinArticle
oldMinArticle=99999999
set rs=conn.execute("select * from usertitle order by MinArticle desc")
do while not rs.eof
	conn.execute("update [user] set userclass='"&rs("usertitle")&"',titlepic='"&rs("titlepic")&"' where (userid>="&request.form("beginid")&" and userid<="&request.form("endid")&") and (article<"&oldMinArticle&" and article>="&rs("MinArticle")&" ) and usergroupid=4")
	oldMinArticle=rs("MinArticle")
	rs.movenext
loop
rs.close
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发帖数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
</form>
<%
end sub

'用户所有主题数
function UserTopicNum(userid)
dim topicnum
topicnum=0
usertopicnum=0
set tmprs=conn.execute("select count(*) from topic where PostUserID="&userid&" and locktopic<2")
TopicNum=tmprs(0)
if isnull(TopicNum) then TopicNum=0
UserTopicNum=UserTopicNum + TopicNum
set tmprs=nothing
end function

'用户所有回复数
function UserReplyNum(userid)
dim replynum
replynum=0
userreplynum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select count(announceid) from "&AllPostTable(i)&" where ParentID>0 and locktopic<2 and PostUserID="&userid)
	replyNum=tmprs(0)
	if isnull(replyNum) then replyNum=0
	UserReplyNum=UserReplyNum + replynum
next
set tmprs=nothing
end function

'用户所有帖子
function Userallnum(userid)
dim allnum
allnum=0
userallnum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select count(announceid) from "&AllPostTable(i)&" where locktopic<2 and  PostUserID="&userid)
	allnum=tmprs(0)
	if isnull(allnum) then allnum=0
	userallnum=userallnum+allnum
next
set tmprs=nothing
end function

'用户所有威望
function UserPowerNum(userid)
dim allPower
allPower=0
userPowerNum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select sum(cint(left(isagree,instr(1,isagree,""|"")-1))-cint(right(isagree,len(isagree)-instr(1,isagree,""|"")))) from "&AllPostTable(i)&" where locktopic<2 and PostUserID="&userid&" and isagree<>''")
	allPower=tmprs(0)
	if isnull(allPower) then allPower=0
	userPowerNum=userPowerNum+allPower
next
set tmprs=nothing
end function

'用户所有积分
function UserScoreNum(userid)
dim allScore
allScore=0
userScoreNum=0
For i=0 to ubound(AllPostTable)
	set tmprs=conn.execute("select sum(Score) from "&AllPostTable(i)&" where locktopic<2 and PostUserID="&userid&" and not isnull(Score)")
	allScore=tmprs(0)
	if isnull(allScore) then allScore=0
	userScoreNum=userScoreNum+allScore
next
set tmprs=nothing
end function
%>