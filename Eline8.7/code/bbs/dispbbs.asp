<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="inc/birthday.asp"-->
<!-- #include file="inc/ubbcode.asp" -->
<!--#include file="z_Wealth.asp"-->
<!--#include file="z_fastpost_conn.asp"-->
<%if Cint(forum_setting(53))=1 and cint(left(forum_setting(72),1))<>2 then%>
<!--#include file="z_Visual_const.asp"-->
<%end if%>
<%
'=========================================================
' File: dispbbs.asp
' Version:5.0
' Date: 2002-9-7
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
Dim AnnounceID
Dim pbtitle
Dim ReplyID
Dim Star,nSkin,SkinPic,Skiname
Dim Topic_1,IsTop,IsBest,IsVote
Dim UserName,view,times
Dim onlineUserList
Dim userhiddensql
Dim Page_Count,TotalRec,abgcolor,bgcolor
Dim TopicCount
Dim Pcount,endpage
Dim isagree,noagree
Dim PostUserName,PostUserid
Dim pollid
Dim TotalUseTable
Dim canreply,mycanreply
Dim LockTopic
dim PostBuyUser
Page_count=0
canreply=false
i=1
if boardmaster or master then
	userhiddensql=""
else
	userhiddensql=" and userhidden=2"
end if

'======================== 这是UBB代码要加入的 =========================

if request("Pay")="yes" and Request.Cookies("eddubb")("PaySee"&request("id")&userid)="yes" then
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>你还有的现金只能让你瞧一眼相关信息了！</li>"
elseif request("Pay")="yes" then
	if mymoney<int(request("PayRmb")) then
		foundErr = true
		if mymoney>=int(request("PayRmb")/2) then
		ErrMsg=ErrMsg+"<br>"+"<li>你还有的现金只能让你瞧一眼相关信息了！</li>"
		else
		ErrMsg=ErrMsg+"<br>"+"<li>你没有足够的现金购买相关信息！</li>"
		end if
	else
		Response.Cookies("eddubb").Expires=Date+999
		Response.Cookies("eddubb")("PaySee"&request("id")&userid) = "yes"
	end if
elseif request("Pay")="one" then
	if mymoney<int(request("PayRmb")/2) then
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>你没有足够的现金购买相关信息！</li>"
	else
		session("Payone"&request("id"))="one"
	end if
end if

' ======================== 到这里结束 =========================

stats="浏览帖子"
if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if request("id")="" then
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
	founderr=true
elseif not isInteger(request("id")) then
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
	founderr=true
else
	AnnounceID=request("id")
	response.cookies("LastView")("Topic_"&AnnounceID&"_"&BoardID)=now()
end if
if request("replyid")="" then
	replyid=Announceid
elseif not isInteger(request("replyid")) then
	replyid=Announceid
else
	replyid=request("replyid")
end if
if request("star")="" or not isnumeric(request("star")) then
	if request("star")="" then
		star=1
	elseif instr(1,request("star"),"#")>0 then
		if split(request("star"),"#")(0)="" or not isnumeric(split(request("star"),"#")(0)) then
			star=1
		else
			star=clng(split(request("star"),"#")(0))
		end if
	end if
else
	star=clng(request("star"))
end if

if request("skin")="" or not isnumeric(request("skin")) then
	skin=Cint(Board_setting(24))
else
	skin=Cint(request("skin"))
end if
if skin=1 then
	nskin=0
	skinpic=Forum_boardpic(12)
	skiname="平板"
elseif skin="0" then
	nskin=1
	skinpic=Forum_boardpic(11)
	skiname="树形"
else
	skin=0
	nskin=1
	skinpic=Forum_boardpic(11)
	skiname="树形"
end if

if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		end if
	end if
end if

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
end if

'###################特殊版面修改开始(asilas制作)##################
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
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
		end if
	end if
end if
'####################特殊版面修改结束(asilas制作)#################

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	conn.execute("update topic set hits=hits+1 where topicid="&Announceid)
	sql="select title,istop,isbest,PostUserName,PostUserid,hits,isvote,child,pollid,LockTopic,PostTable from topic where topicID="&Announceid
	set rs=conn.execute(sql)
	if not(rs.bof and rs.eof) then
		if rs("locktopic")=2 then
			ErrMsg=ErrMsg+"<br>"+"<li>该帖子已经被管理员删除！</li>"
			founderr=true
		end if
		topic_1=rs(0)
		istop=rs(1)
		isbest=rs(2)
		PostUserName=rs(3)
		PostUserID=rs(4)
		view=rs(5)
		isVote=rs(6)
		TopicCount=rs(7)+1
		pollid=rs(8)
		Locktopic=rs(9)
		TotalUseTable=rs(10)
		stats=topic_1
		if PostUserName=membername then
			call readRe()
			mycanreply=GroupSetting(4)
		else
			mycanreply=GroupSetting(5)
			if Cint(GroupSetting(2))=0 then
				Errmsg=Errmsg+"<br>"+"<li>您没有浏览在本论坛查看其他人发布的帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
				founderr=true
			end if
		end if
	else
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的帖子不存在</li>"
		founderr=true
	end if
	set rs=nothing
	call nav()
	call head_var(1,BoardDepth,0,0)
	if founderr then
		call dvbbs_error()
	else
		call main()
		if founderr then call dvbbs_error()
	end if
end if
call footer()

sub main()%>
<SCRIPT LANGUAGE="JavaScript">
var isRunning=0;
var Rate=0;
var delaWidth=5;
function ResizeImg()
{
	if(!isRunning && document.all.ImgSpan!=null) {
		isRunning=1;
		var ImgSpanArray = ImgSpan;
		var i,j,OldWidth;
		var containerobj;
		
		if(ImgSpanArray.length)
			for(i=0;i<ImgSpanArray.length;i++) {
				for(j=0;j<ImgSpanArray[i].children.length;j++)
					if(ImgSpanArray[i].children[j].tagName=="IMG") {
						containerobj=ImgSpanArray[i].children[j].parentElement;
						while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
						if(ImgSpanArray[i].children[j].width>containerobj.clientWidth-delaWidth) {
							ImgSpanArray[i].children[j].width=containerobj.clientWidth-delaWidth;
						}
						else {
							OldWidth=ImgSpanArray[i].children[j].className;
							if(ImgSpanArray[i].children[j].width<OldWidth) {
								containerobj=ImgSpanArray[i].children[j].parentElement;
								while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
								if(OldWidth>containerobj.clientWidth-delaWidth) ImgSpanArray[i].children[j].width=containerobj.clientWidth-delaWidth;
								else ImgSpanArray[i].children[j].width=OldWidth;
							}
						}
					}
			}
		else
			for(j=0;j<ImgSpanArray.children.length;j++)
				if(ImgSpanArray.children[j].tagName=="IMG") {
					containerobj=ImgSpanArray.children[j].parentElement;
					while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
					if(ImgSpanArray.children[j].width>containerobj.clientWidth-delaWidth) {
						ImgSpanArray.children[j].width=containerobj.clientWidth-delaWidth;
					}
					else {
						OldWidth=ImgSpanArray.children[j].className;
						if(ImgSpanArray.children[j].width<OldWidth) {
							containerobj=ImgSpanArray.children[j].parentElement;
							while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
							if(OldWidth>containerobj.clientWidth-delaWidth) ImgSpanArray.children[j].width=containerobj.clientWidth-delaWidth;
							else ImgSpanArray.children[j].width=OldWidth;
						}
					}
				}
		isRunning=0;
	}
}
function LoadImg()
{
	isRunning=1;
	if(document.all.ImgSpan!=null) {
		var ImgSpanArray = ImgSpan;
		var i,j,OldWidth;
		var containerobj;
		
		if(ImgSpanArray.length)
			for(i=0;i<ImgSpanArray.length;i++) {
				for(j=0;j<ImgSpanArray[i].children.length;j++) 
					if(ImgSpanArray[i].children[j].tagName=="IMG") {
						if(ImgSpanArray[i].children[j].complete) {
							ImgSpanArray[i].children[j].className=ImgSpanArray[i].children[j].width;
							containerobj=ImgSpanArray[i].children[j].parentElement;
							while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
							if(ImgSpanArray[i].children[j].width>containerobj.clientWidth-delaWidth) {
								ImgSpanArray[i].children[j].width=containerobj.clientWidth-delaWidth;
							}
						}
					}
			}
		else
			for(j=0;j<ImgSpanArray.children.length;j++)
				if(ImgSpanArray.children[j].tagName=="IMG") {
					if(ImgSpanArray.children[j].complete) {
						ImgSpanArray.children[j].className=ImgSpanArray.children[j].width;
						containerobj=ImgSpanArray.children[j].parentElement;
						while (containerobj.tagName!="TD"||containerobj.name!="bbscontent") containerobj=containerobj.parentElement;
						if(ImgSpanArray.children[j].width>containerobj.clientWidth-delaWidth) {
							ImgSpanArray.children[j].width=containerobj.clientWidth-delaWidth;
						}
					}
				}
	}
	isRunning=0;
}

window.onload = LoadImg;
window.onresize = ResizeImg;
</script>
<%if ((not cint(Board_Setting(0))=1) and Cint(mycanreply)=1 and Cint(locktopic)=0) or (master or superboardmaster or boardmaster) then
	canreply=true
	select case cint(forum_setting(73))
	case 2
		if myUserPPD*cint(forum_setting(74))<=myTodayReply then
			canreply=false
		end if
	case 3
		if myUserPPD*cint(forum_setting(74))<=myTodayTopic+myTodayReply then
			canreply=false
		end if
	end select
end if

if IsVote=1 then
	Dim vrs,vote,vote_1,votenum,votenum_1,m,g,votetype
	Dim vurs,voteyes
	voteyes=0
	g=0
	sql="select * from vote where VoteID="&PollID
	set vrs=server.createobject("adodb.recordset")
	vrs.open sql,conn,1,1
	if not (vrs.eof and vrs.bof) then
	voteyes=1
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style=""table-layout:fixed;word-break:break-all""><tr><th align=left colspan=2 height=25>[投票]："&Topic_1&"</th></tr><form action=postvote.asp?BoardID="&BoardID&"&voteID="&PollID&"&id="&Announceid&"&action="&vrs("votetype")&" method=POST>"
	vote=split(vrs("vote"),"|")
	votenum=split(vrs("votenum"),"|")
	for i = 0 to ubound(votenum)
		votenum_1=cint(votenum_1)+votenum(i)
	next
	if votenum_1=0 then votenum_1=1
	for m = 0 to ubound(vote)
		g=g+1
		if g=11 then g=1
		if cint(vrs("votetype"))=0 then
		votetype="<input type=radio name=postvote value="""&m&""">"
		else
		votetype="<input type=checkbox name=postvote_"&m&" value="""&m&""">"
		end if
		response.write "<tr><td width=""60%"" height=25 class=tablebody1>"&m+1&".  "&votetype & htmlencode(vote(m))&"</td><td width=""40%"" class=tablebody1><img src="""&Forum_info(7)&"bar"&g&".gif"" width="""&Cint(replace(FormatPercent(votenum(m)/votenum_1),"%",""))*3.3&""" height=8> <b>"&votenum(m)&"票</b></td></tr>"
	next

	if not founduser or datediff("d",vrs("timeout"),Now())>0 then
		response.write "<tr><td class=tablebody2 colspan=2 height=25>您还没有登录，不能进行投票；或者已经过了投票期限。[<a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
	else
		set vurs=conn.execute("select userid from voteuser where voteid="&PollID&" and userid="&userid)
		if vurs.eof and vurs.bof then
			response.write "<tr><td colspan=2 height=25 class=tablebody1><input type=submit name=Submit value='投 票'>  [截止时间："&vrs("timeout")&" | <a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
		else
			response.write "<tr><td class=tablebody2 colspan=2 height=25>您已经投过票了，请看结果吧。[过期时间："&vrs("timeout")&" | <a href=javascript:openScript('viewvoters.asp?id="&pollid&"',300,500)>查看投票用户</a>]</td></tr>"
		end if
		set vurs=nothing
	end if
	response.write "</form></table><BR>"
	end if
	vrs.close
	set vrs=nothing
end if
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center>
	<tr>
	<td align=left width="40%" valign=middle>&nbsp; 
<%if (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=0) or (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=1 and (master or superboardmaster or boardmaster)) then
	select case cint(forum_setting(73))
	case 1
		if myUserPPD*cint(forum_setting(74))>myTodayTopic then
			response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""发表一个新主题"" border=0></a>&nbsp;"
			response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""发表一个新投票"" border=0></a>&nbsp;"
		else
			response.write "<B>发帖数到达限制，不允许进行发帖操作</B>"
		end if
		response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&Forum_boardpic(4)&""" alt=""回复主题"" border=0></a>"
	case 2
		response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""发表一个新主题"" border=0></a>&nbsp;"
		response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""发表一个新投票"" border=0></a>&nbsp;"
		if myUserPPD*cint(forum_setting(74))>myTodayReply then
			response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&Forum_boardpic(4)&""" alt=""回复主题"" border=0></a>"
		else
			response.write "<B>回复数到达限制，不允许进行回帖操作</B>"
		end if
	case 3
		if myUserPPD*cint(forum_setting(74))>myTodayTopic+myTodayReply then
			response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""发表一个新主题"" border=0></a>&nbsp;"
			response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""发表一个新投票"" border=0></a>&nbsp;"
			response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&Forum_boardpic(4)&""" alt=""回复主题"" border=0></a>"
		else
			response.write "<B>帖子数到达限制，不允许进行发帖/回帖等操作</B>"
		end if
	case else
		response.write "<a href=""announce.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(1)&""" alt=""发表一个新主题"" border=0></a>&nbsp;"
		response.write "<a href=""vote.asp?BoardID="&BoardID&"""><img src="""&Forum_info(7)&Forum_boardpic(2)&""" alt=""发表一个新投票"" border=0></a>&nbsp;"
		response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&Forum_boardpic(4)&""" alt=""回复主题"" border=0></a>"
	end select
else
	response.write "<B>本论坛已锁定，不允许进行发帖/回帖等操作</B>"
end if%>
	</td>
	<td align=right width="60%" valign=middle>您是本帖的第 <B><%=view%></B> 个阅读者<a href="go.asp?BoardID=<%=BoardID%>&sid=<%=Announceid%>"><img src="<%=Forum_info(7)&Forum_boardpic(14)%>" border=0 alt=浏览上一篇主题 width=52 height=12></a>&nbsp;
	<a href="javascript:this.location.reload()"><img src="<%=Forum_info(7)&Forum_statePic(7)%>" border=0 alt=刷新本主题 width=40 height=12></a> &nbsp;
	<a href="?BoardID=<%=BoardID%>&replyID=<%=replyID%>&id=<%=request("id")%>&star=<%=star%>&skin=<%=nskin%>"><img src="<%=Forum_info(7)&skinpic%>" width=40 height=12 border=0 alt=<%=skiname%>显示帖子></a>　<a href="go.asp?BoardID=<%=BoardID%>&sid=<%=Announceid%>&action=next"><img src="<%=Forum_info(7)&Forum_boardpic(13)%>" border=0 alt=浏览下一篇主题 width=52 height=12></a>
	</td>
	</tr>
</table>

<TABLE cellPadding=0 cellSpacing=1 align=center class=tableborder1>
	<tr align=middle> 
	<td align=left valign=middle width="100%" height=25>
		<table width=100% cellPadding=0 cellSpacing=0 border=0>
		<tr>
		<th align=left valign=middle width="73%" height=25>
		&nbsp;* 帖子主题</B>： <%=htmlencode(topic_1)%></th>
		<th width=37% align=right>
		<a href="report.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src=<%=Forum_info(7)&Forum_TopicPic(1)%> alt=报告本帖给版主 border=0></a>&nbsp;
		<a href="printpage.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(2)%>" alt=显示可打印的版本 border=0></a>&nbsp;
		<a href="pag.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(3)%>" border=0 alt=把本帖打包邮递></a>&nbsp;
		<a href="favadd.asp?BoardID=<%=BoardID%>&id=<%=Announceid%>"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(4)%>" BORDER=0 alt=把本帖加入论坛收藏夹></a>&nbsp;
		<a href="sendpage.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>"><img src="<%=Forum_info(7)&Forum_TopicPic(5)%>" border=0 alt=发送本页面给朋友></a>&nbsp;
		<a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>/dispbbs.asp?BoardID=<%=BoardID%>&id=<%=AnnounceID%>', ' <%=Forum_info(0)%> - <%=htmlencode(replace(topic_1,"&#39;",""))%>')"><IMG SRC="<%=Forum_info(7)&Forum_TopicPic(6)%>" BORDER=0 width=15 height=15 alt=把本帖加入IE收藏夹></a>&nbsp;
		</th>
		</tr>
		</table>
	</td>
	</tr>
</table>

<%
if request("action")="dispaudit" then
	dim AdminLockTopic
	AdminLockTopic=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(36))=1 then
		AdminLockTopic=true
	else
		AdminLockTopic=false
	end if
	if Cint(GroupSetting(36))=1 and UserGroupID>3 then
		AdminLockTopic=true
	end if
	if FoundUserPer and Cint(GroupSetting(36))=1 then
		AdminLockTopic=true
	elseif FoundUserPer and Cint(GroupSetting(36))=0 then
		AdminLockTopic=false
	end if
	if not AdminLockTopic then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本版面审核帖子的权限。"
		founderr=true
		exit sub
	end if
	sql="Select B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.VIP,B.Speak,b.score,b.scoreuser,b.scoretime,u.UserScore,u.wife from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.AnnounceID="&replyID
else
if cint(skin)=1 and replyid=Announceid then
	sql="Select top 1  B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.VIP,B.Speak,b.score,b.scoreuser,b.scoretime,u.UserScore,u.wife from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.rootID="&AnnounceID&" and B.locktopic<2 order by Announceid"
elseif cint(skin)=1 then
	sql="Select B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.VIP,B.Speak,b.score,b.scoreuser,b.scoretime,u.UserScore,u.wife from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.AnnounceID="&replyID&" and  B.locktopic<2"
else
	sql="Select B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.homepage,U.oicq,U.sign,U.userclass,U.title,U.width,U.height,U.article,U.face,U.addDate,U.userWealth,U.userEP,U.userCP,U.birthday,U.sex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.VIP,B.Speak,b.score,b.scoreuser,b.scoretime,u.UserScore,u.wife from "&TotalUseTable&" B inner join [user] U on U.UserID=B.PostUserID where B.BoardID="&BoardID&" and B.RootID="&Announceid&" and B.locktopic<2 order by B.AnnounceID"
end if
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>该帖子无法浏览，可能的原因有：帖子已被删除、发帖用户不存在、该帖子正在审核中。"
	Founderr=true
	exit sub
end if
if not(rs.eof and rs.bof) then
	if TopicCount mod Cint(Board_Setting(27))=0 then
		Pcount= TopicCount \ Cint(Board_Setting(27))
	else
		Pcount= TopicCount \ Cint(Board_Setting(27))+1
	end if
	if star > Pcount then star = Pcount
  if star < 1 then star = 1
	rs.PageSize = Cint(Board_Setting(27))
	rs.AbsolutePage=star
	page_count=0
	Dim useremail,homepage,oicq,sign,userclass,title,width,uheight,article,face
	Dim addtime,userWealth,userEP,userCP,userbirth,sex,UserGroup,usercolor,namestyle,UserScore
	Dim LockUser,userPower,istopic,PostUserGroup,UserLastLogin,userip,reUserGroupID
	Dim bestTitle,istoptitle,newbbs
	Dim RootID,signflag,isbest,PostUserID,layer,isagree,bbsisagree
	dim srs
	istopic=false
'AnnounceID=0,BoardID=1,UserName=2,Topic=3,dateandtime=4,body=5,Expression=6,ip=7,RootID=8,signflag=9,isbest=10,PostUserid=11,layer=12,isagree=13,useremail=14,homepage=15,oicq=16,sign=17,userclass=18,title=19,width=20,height=21,article=22,face=23,addDate=24,userWealth=25,userEP=26,userCP=27,birthday=28,sex=29,UserGroup=30,LockUser=31,userPower=32,titlepic=33,UserGroupID=34,LastLogin=35
	if cint(forum_setting(27))=1 then newbbs=CurNewTopic(20,10,3)
	do while not rs.eof and page_count < Cint(Board_Setting(27))
	username=rs(2)
	RootID=rs(8)
	signflag=rs(9)
	isbest=rs(10)
	PostUserID=rs(11)
	layer=rs(12)
	isagree=rs(13)
	if isnull(isagree) or isagree="" then isagree="0|0"
	bbsisagree=split(isagree,"|")
	replyid=rs(0)
	if star=1 and skin=0 and layer=1 then
		istopic=true
		noagree=bbsisagree(1)
		isagree=bbsisagree(0)
	elseif layer=1 and skin=1 then
		istopic=true
		noagree=bbsisagree(1)
		isagree=bbsisagree(0)
	end if
	useremail=rs(14)
	homepage=rs(15)
	oicq=rs(16)
	sign=trim(rs(17))
	userclass=rs(18)
	title=rs(19)
	width=rs(20)
	if isnull(width) or width=0 then
		width="width="&forum_setting(38)
	else
		width="width="&width
	end if
	uheight=rs(21)
	if isnull(uheight) or uheight=0 then
		uheight="height="&forum_setting(39)
	else
		uheight="height="&uheight
	end if
	article=rs(22)
	face=htmlencode(FilterJS(rs(23)))
	addtime=rs(24)
	if not isdate(addtime) then addtime=now()
	userWealth=rs(25)
	userEP=rs(26)
	userCP=rs(27)
	userbirth=rs(28)
	sex=rs(29)
	UserGroup=rs(30)
	LockUser=rs(31)
	userPower=rs(32)
	reUserGroupID=rs(34)
	UserLastLogin=rs(35)
	PostBuyUser=rs(36)
	userScore=rs("UserScore")
	if reUserGroupID=8 then
	namestyle="filter:glow(color="&Forum_body(22)&",strength=2)"
	usercolor=Forum_body(21)
	elseif reUserGroupID=3 then
	namestyle="filter:glow(color="&Forum_body(18)&",strength=2)"
	usercolor=Forum_body(17)
	elseif reUserGroupID=1 then
	namestyle="filter:glow(color="&Forum_body(20)&",strength=2)"
	usercolor=Forum_body(19)
	else
	namestyle="filter:glow(color="&Forum_body(16)&",strength=2)"
	usercolor=Forum_body(15)
	end if
	if bgcolor="tablebody1" then
		bgcolor="tablebody2"
		abgcolor="tablebody1"
	else
		bgcolor="tablebody1"
		abgcolor="tablebody2"
	end if
	if userbirth="" or isnull(userbirth) then
	userbirth=""
	else
	userbirth=astro(userbirth)
	end if
	response.write "<TABLE cellPadding=5 cellSpacing=1 align=center class=tableborder1 style="" table-layout:fixed;word-break:break-all"">"
	response.write "<tr><td class="&bgcolor&" width=175 valign=top>"
	response.write "<table width=100% cellpadding=4 cellspacing=0><tr>"
	response.write "<td width=* valign=middle style="""&namestyle&""" nowrap>&nbsp;<a name="&rs("AnnounceID")&"><font color="&usercolor&"><B>"&htmlencode(UserName)&"</B></font></a>	</td>"
	response.write "<td width=40 valign=middle>"&userbirth&"</td>"
	response.write "</tr></table>"
	if Cint(forum_setting(53))=1 and face<>"" then 
		if cint(left(forum_setting(72),1))=0 then
			call ShowUserVisual(UserName,140,bgcolor,false)
		elseif cint(left(forum_setting(72),1))=1 then
			call ShowUserVisual(UserName,140,bgcolor,true)
		else
			response.write "&nbsp;&nbsp;<img src="&htmlencode(FilterJS(face))&" "&width&" "&uheight&">"
		end if
	end if
	if face="" then
	response.write "&nbsp;&nbsp;<img src=noface.gif>"
        end if
	response.write "<br>"
	if cint(right(forum_setting(72),1))<>1 then
		response.write "&nbsp;&nbsp;<img src="&Forum_info(7)&rs("titlepic")&"><br>"
	else
		response.write "&nbsp;&nbsp;<img src="&Forum_info(7)&rs("titlepic")&" alt="""
	end if
	if cint(right(forum_setting(72),1))<>1 then
		if rs("vip")=1 then response.write "&nbsp;&nbsp;属性：<font color='" & forum_body(8) & "'><B>VIP会员</b></font><BR>"
		if title<>"" then response.write "&nbsp;&nbsp;头衔：" & htmlencode(title) & "<br>"
		response.write "&nbsp;&nbsp;等级："& userclass &"<BR>"
		if UserGroup<>"" then
		response.write "&nbsp;&nbsp;门派：" & UserGroup
		if conn.execute("select GroupName from [GroupName] where Zangmen='"&username&"'").eof then
			response.write "(弟子)"
		else
			response.write "(掌门)"
		end if
		response.write "<br>"
		end if
  end if
  if userScore > 0 then 
		response.write "&nbsp;&nbsp;积分：" & "<font color='"&forum_body(8)&"'><strong>+"&userScore&"</strong></font>"&"&nbsp;&nbsp;<br>"
	else
		response.write "&nbsp;&nbsp;积分：" & "<font color='"&forum_body(8)&"'><strong>"&userScore&"</strong></font>"&"&nbsp;&nbsp;<br>"
	end if
  if userPower > 0 then 
  	response.write "&nbsp;&nbsp;威望：" 
  	response.Write "<font color='blue'><strong>+"&userPower&"</strong></font>&nbsp;&nbsp;<br>" 
  else 
  	response.write "&nbsp;&nbsp;威望：" 
  	response.Write "<font color='blue'><strong>"&userPower&"</strong></font>&nbsp;&nbsp;<br>" 
  end if 
	response.write "&nbsp;&nbsp;魅力："&userCP&"&nbsp;&nbsp;<br>"
	response.write "&nbsp;&nbsp;经验："&userEP&"&nbsp;&nbsp;<br>"
	response.write "&nbsp;&nbsp;现金："&userwealth&"&nbsp;&nbsp;<br>"
	if cint(forum_Setting(45))=1 then
		response.write "&nbsp;&nbsp;存款："&bank_money(username)&"&nbsp;&nbsp;<br>"
		response.write "&nbsp;&nbsp;股票："&stock_money(username)&"&nbsp;&nbsp;<br>"
		response.write "&nbsp;&nbsp;资产："&(userwealth+bank_money(username)+stock_money(username))&"&nbsp;&nbsp;<br>"
	end if
	response.write "&nbsp;&nbsp;文章："& article & "&nbsp;&nbsp;<BR>"
if not isnull(rs("wife")) or not trim(rs("wife"))="" then
		if rs("sex")=1 then
			response.write "&nbsp;&nbsp;配偶：<a href=dispuser.asp?name="&htmlencode(trim(rs("wife")))&"><font color=red>"&htmlencode(trim(rs("wife")))&"</font></a><br>"
		else
			response.write "&nbsp;&nbsp;配偶：<a href=dispuser.asp?name="&htmlencode(trim(rs("wife")))&"><font color=blue>"&htmlencode(trim(rs("wife")))&"</font></a><br>"
		end if
	end if
	if cint(right(forum_setting(72),1))=1 then
		response.write """><br>"
		if rs("vip")=1 then response.write "&nbsp;&nbsp;属性：<font color='" & forum_body(8) & "'><B>VIP会员</b></font><BR>"
		if title<>"" then response.write "&nbsp;&nbsp;头衔：" & htmlencode(title) & "<br>"
		response.write "&nbsp;&nbsp;等级："& userclass &"<BR>"
		if UserGroup<>"" then
		response.write "&nbsp;&nbsp;门派：" & UserGroup
		if conn.execute("select GroupName from [GroupName] where Zangmen='"&username&"'").eof then
			response.write "(弟子)"
		else
			response.write "(掌门)"
		end if
		response.write "<br>"
		end if
  end if
	response.write "<br>"
	if boardmaster or master or superboardmaster then
		set Srs=conn.execute("select o.boardid,b.boardtype from online o left join board b on o.boardid=b.boardid where o.username='"&username&"'")
	else
		set Srs=conn.execute("select o.boardid,b.boardtype from online o left join board b on o.boardid=b.boardid where o.username='"&username&"' and o.userhidden=2")
	end if
	if Srs.eof and Srs.bof then
		response.write "&nbsp;&nbsp;状态：离线"
	else
		if isnull(srs(1)) then
			response.write "&nbsp;&nbsp;状态：论坛首页"
		else	
			response.write "&nbsp;&nbsp;状态：<a href=list.asp?boardid="&srs(0)&" target=_blank><font color=blue>"&srs(1)&"</font></a>"
		end if
	end if
	Srs.close
	set Srs=nothing
	response.write "<br>"
	response.write "&nbsp;&nbsp;注册："& year(addtime) &"-"& month(addtime) &"-"& day(addtime)
	response.write "<br>"
	if istopic then 
		response.write "<br>&nbsp;&nbsp;<a href=""postagree.asp?boardid="&boardid&"&id="&Announceid&"&userid="&PostUserID&"&isagree=1"" title=""同意该帖观点，给他一朵鲜花，将消耗您"&GroupSetting(47)&"点金钱""><img src="&Forum_info(7)&Forum_BoardPic(6)&" border=0>鲜花</a>(<font color="&Forum_body(8)&">"&isagree&"</font>)&nbsp;&nbsp;<a href=""postagree.asp?boardid="&boardid&"&id="&Announceid&"&userid="&PostUserID&"&isagree=2"" title=""不同意该帖观点，给他一个鸡蛋，将消耗您"&GroupSetting(47)&"点金钱""><img src="&Forum_info(7)&Forum_BoardPic(7)&" border=0>鸡蛋</a>(<font color="&Forum_body(8)&">"&noagree&"</font>)<br>"
	end if

response.write "<br>"
response.write "&nbsp;&nbsp;<a href=z_jifen.asp><font color=#cc0000>╭☆╯保存积分先╭☆╯</font></a><br><br>"
response.write "&nbsp;&nbsp;<a href=z_dglistall.asp><font color=blue>╭☆╯点歌送好友╭☆╯</font></a><br><br>"
response.write "&nbsp;&nbsp;<a href=z_visual.asp><font color=#aa00ff>╭☆╯我要扮靓靓╭☆╯</font></a><br><br>"
response.write "&nbsp;&nbsp;<a href=z_visual.asp?shopid=197><font color=#ff0000>╭☆╯靚靚照相館╭☆╯</font></a><br>"

	response.write "</td><td class="&bgcolor&" valign=top width=* height=""100%"">"
	response.write "<table width=100%  height=""100%"">"
	response.write "<tr>"
	response.write "<script>"
	response.write "var status0_"&rs(0)&"='"
	response.write "<a href=\""javascript:openScript(\'messanger.asp?action=new&touser="&HTMLEncode(UserName)&"\',500,400)\""><img src=\"""&Forum_info(7)&Forum_TopicPic(7)&"\"" border=0 alt=\""给"&HTMLEncode(UserName)&"发送一个短消息\""></a>&nbsp;"
	response.write "<a href=\""friendlist.asp?action=addF&myFriend="&HTMLEncode(UserName)&"\"" target=_blank><img src=\"""&Forum_info(7)&Forum_TopicPic(21)&"\"" border=0 alt=\""把"&HTMLEncode(UserName)&"加入好友\""></a>&nbsp;"
	response.write "<a href=\""dispuser.asp?id="&postuserid&"\"" target=_blank><img src=\"""&Forum_info(7)&Forum_TopicPic(9)&"\"" border=0 alt=\""查看"&HTMLEncode(UserName)&"的个人资料\""></a>&nbsp;"
	response.write "<a href=\""queryResult.asp?stype=1&nSearch=3&keyword="&HTMLEncode(UserName)&"&BoardID="&cstr(BoardID)&"&SearchDate=ALL\"" target=_blank><img src=\"""&Forum_info(7)&Forum_TopicPic(8)&"\"" border=0 alt=\""搜索"&HTMLEncode(UserName)&"在"&boardtype&"的所有帖子\""></a>&nbsp;"
	response.write "<a href=\""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"&reply=true\""><img src=\"""&Forum_info(7)&Forum_TopicPic(15)&"\"" border=0 alt=\""引用回复这个帖子\""></a>&nbsp;"
	response.write "<a href=\""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"\""><img src=\"""&Forum_info(7)&Forum_TopicPic(22)&"\"" border=0 alt=\""回复这个帖子\""></a>&nbsp;"
	response.write "<img src="&Forum_info(7)&Forum_boardpic(16)&" border=0 style=\""cursor:hand;position:relative;top:-3;\"" onclick=\'if(curfontsize_"&rs(0)&">8){fontsize_"&rs(0)&".style.fontSize=(--curfontsize_"&rs(0)&")+\""pt\"";fontsize_"&rs(0)&".style.lineHeight=(--curlineheight_"&rs(0)&")+\""pt\"";}\'><font style=\""position:relative;top:-3;\"">调整字体</font><img src="&Forum_info(7)&Forum_boardpic(15)&" border=0 style=\""cursor:hand;position:relative;top:-2;\"" onclick=\'if(curfontsize_"&rs(0)&"<64){fontsize_"&rs(0)&".style.fontSize=(++curfontsize_"&rs(0)&")+\""pt\"";fontsize_"&rs(0)&".style.lineHeight=(++curlineheight_"&rs(0)&")+\""pt\"";}\'>&nbsp;"
	response.write "<img src=\"""&Forum_info(7)&Forum_pic(12)&"\"" align=absmiddle alt=\""其他\"" style=\""cursor:hand\"" onclick=\""if(statusflag_"&rs(0)&"==0) this.parentElement.innerHTML=status1_"&rs(0)&";else this.parentElement.innerHTML=status0_"&rs(0)&";statusflag_"&rs(0)&"=1-statusflag_"&rs(0)&"\"">"
	response.write "';"
	response.write "var status1_"&rs(0)&"='"
	'去掉点歌
	'response.write "<a href=\""z_dgwrite.asp?name="&HTMLEncode(UserName)&"\"" target=_blank><img src=pic/dg.gif border=0 alt=\""为"&HTMLEncode(UserName)&"点歌祝福\""></a>&nbsp;"
	if useremail<>"" then response.write "<A href=\""mailto:"& htmlencode(useremail) &"\""><IMG alt=\""点击这里发送电邮给"& HTMLEncode(UserName) &"\"" border=0 src="&Forum_info(7)&Forum_TopicPic(10)&"></A>&nbsp;"
	if oicq<>"" then response.write "<a href=\""http://search.tencent.com/cgi-bin/friend/user_show_info?ln="&oicq&"\"" target=_blank><img src="&Forum_info(7)&Forum_TopicPic(11)&" border=0 alt=\"""&htmlencode(UserName)&"["&oicq&"]的QQ情况\""></a>&nbsp;"
	if homepage<>"" then response.write "<A href=\"""& htmlencode(homepage) &"\"" target=_blank><IMG alt=\""访问"& HTMLEncode(UserName) &"的主页\"" border=0 src="&Forum_info(7)&Forum_TopicPic(14)&"></A>&nbsp;"
	response.write "<img src=\"""&Forum_info(7)&Forum_pic(12)&"\"" align=absmiddle alt=\""其他\"" style=\""cursor:hand\"" onclick=\""if(statusflag_"&rs(0)&"==0) this.parentElement.innerHTML=status1_"&rs(0)&";else this.parentElement.innerHTML=status0_"&rs(0)&";statusflag_"&rs(0)&"=1-statusflag_"&rs(0)&"\"">"
	response.write "';"
	response.write "var statusflag_"&rs(0)&"=0;"
	response.write "var curfontsize_"&rs(0)&"="&board_setting(28)&";"
	response.write "var curlineheight_"&rs(0)&"="&board_setting(29)&";"
	response.write "</script>"
	response.write "<td width=420 height=""20"" nowrap>"
	response.write "<a href=""javascript:openScript('messanger.asp?action=new&touser="&HTMLEncode(UserName)&"',500,400)""><img src="""&Forum_info(7)&Forum_TopicPic(7)&""" border=0 alt=""给"&HTMLEncode(UserName)&"发送一个短消息""></a>&nbsp;"
	response.write "<a href=""friendlist.asp?action=addF&myFriend="&HTMLEncode(UserName)&""" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(21)&""" border=0 alt=""把"&HTMLEncode(UserName)&"加入好友""></a>&nbsp;"
	response.write "<a href=""dispuser.asp?id="&postuserid&""" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(9)&""" border=0 alt=""查看"&HTMLEncode(UserName)&"的个人资料""></a>&nbsp;"
	response.write "<a href=""queryResult.asp?stype=1&nSearch=3&keyword="&HTMLEncode(UserName)&"&BoardID="&cstr(BoardID)&"&SearchDate=ALL"" target=_blank><img src="""&Forum_info(7)&Forum_TopicPic(8)&""" border=0 alt=""搜索"&HTMLEncode(UserName)&"在"&boardtype&"的所有帖子""></a>&nbsp;"
	response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"&reply=true""><img src="""&Forum_info(7)&Forum_TopicPic(15)&""" border=0 alt=引用回复这个帖子></a>&nbsp;"
	response.write "<a href=""reannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"""><img src="""&Forum_info(7)&Forum_TopicPic(22)&""" border=0 alt=回复这个帖子></a>&nbsp;"
	response.write "<img src="&Forum_info(7)&Forum_boardpic(16)&" border=0 style=""cursor:hand;position:relative;top:-3;"" onclick='if(curfontsize_"&rs(0)&">8){fontsize_"&rs(0)&".style.fontSize=(--curfontsize_"&rs(0)&")+""pt"";fontsize_"&rs(0)&".style.lineHeight=(--curlineheight_"&rs(0)&")+""pt"";}'><font style=""position:relative;top:-3;"">调整字体</font><img src="&Forum_info(7)&Forum_boardpic(15)&" border=0 style=""cursor:hand;position:relative;top:-2;"" onclick='if(curfontsize_"&rs(0)&"<64){fontsize_"&rs(0)&".style.fontSize=(++curfontsize_"&rs(0)&")+""pt"";fontsize_"&rs(0)&".style.lineHeight=(++curlineheight_"&rs(0)&")+""pt"";}'>&nbsp;"
	response.write "<img src="""&Forum_info(7)&Forum_pic(12)&""" align=absmiddle alt=""其他"" style=""cursor:hand"" onclick=""if(statusflag_"&rs(0)&"==0) this.parentElement.innerHTML=status1_"&rs(0)&";else this.parentElement.innerHTML=status0_"&rs(0)&";statusflag_"&rs(0)&"=1-statusflag_"&rs(0)&""">"
	response.write "</td><td width=* nowrap align=right>"
	dim CanScore
	if cint(board_Setting(49))=0 then
		CanScore=1
	else
		CanScore=isCanScore(rs("scoreuser"),username)
	end if
	if CanScore=0 then
		if isnull(rs("scoreuser")) or not (rs("scoreuser")<>"") then
			response.write "<script>var rate_"&rs(0)&"=0;</script>"
		else
			response.write "<script>var rate_"&rs(0)&"="&rs("score")&";</script>"
		end if
		response.write "<img src='images/img_rate/left.gif' style='CURSOR:hand' onclick='if(rate_"&rs(0)&">-5){rate_"&rs(0)&"=(rate_"&rs(0)&"+4)%11-5;rateimage_"&rs(0)&".src=""images/img_rate/""+rate_"&rs(0)&"+"".gif"";}'>"
		response.write "<a href=""z_score_post.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"&rate="" onclick='this.href+=rate_"&rs(0)&"'>"
		if isnull(rs("scoreuser")) or not (rs("scoreuser")<>"") then
			response.write "<img src='images/img_rate/0.gif' id='rateimage_"&rs(0)&"' border=0 alt=""给这个帖子评分"">"
		else
			response.write "<img src='images/img_rate/"&rs("score")&".gif' id='rateimage_"&rs(0)&"' border=0 alt=""评 分 人："&rs("scoreuser")&"<br>评分时间："&formatdatetime(rs("scoretime"),2)&""">"
		end if
		response.write "</a>"
		response.write "<img src='images/img_rate/right.gif' style='CURSOR:hand' onclick='if(rate_"&rs(0)&"<5){rate_"&rs(0)&"=(rate_"&rs(0)&"+6)%11-5;rateimage_"&rs(0)&".src=""images/img_rate/""+rate_"&rs(0)&"+"".gif"";}'>"
	elseif cint(board_Setting(49))=1 then
		if isnull(rs("scoreuser")) or not (rs("scoreuser")<>"") then
			response.write "<img src='images/img_rate/0.gif' id='rateimage_"&rs(0)&"' border=0 alt=""该帖尚未评分"">"
		else
			response.write "<img src='images/img_rate/"&rs("score")&".gif' id='rateimage_"&rs(0)&"' border=0 alt=""评 分 人："&rs("scoreuser")&"<br>评分时间："&formatdatetime(rs("scoretime"),2)&""">"
		end if
	end if
	response.write "&nbsp;&nbsp;<b>"
	if star=1 and page_count=0 then
	response.write "楼主"
	else
	response.write "第<font color="&Forum_body(8)&">" &(star-1)*Cint(Board_Setting(27))+page_count+1 & "</font>楼"
	end if
	response.write "</b></td></tr><tr><td bgcolor="&Forum_body(27)&" height=1 colspan=2></td></tr>"
	response.write "<tr><td colspan=2 height=1>&nbsp;</td></tr>"
	response.write "<tr><td colspan=2 valign=top><table class=tablebody2 border=0 width=""100%"" style=""table-layout:fixed;word-break:break-all"" cellspacing=""0"" cellpadding=""0"">"
	response.write "<tr>"
	response.write "<td width=20></td>"
	response.write "<td background=""images/haha/top_l.gif"" width=""14"" height=""8""></td>"
	response.write "<td background=""images/haha/top_c.gif""></td>"
	response.write "<td background=""images/haha/top_r.gif"" width=""16""></td>"
	response.write "<td width=20></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td width=20></td>"

	'帖子加挂件
	
	response.write "<td background=""images/haha/center_l.gif"" width=""14"" ></td>"
        response.write "<td><table class=tablebody2 border=0 cellspacing=0 cellpadding=0 style="" table-layout:fixed;word-break:break-all""><tr> <td valign=""top"" width=""85%"" style=""font-size:"&Board_setting(28)&"pt;line-height:"&Board_setting(29)&"pt;WORD-WRAP: break-word;"" bgColor=#fffff1 name=""bbscontent"">"
        if LockUser=2 then
        response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>该用户发言已被管理员屏蔽</font><br>================================</DIV>"
        elseif rs("isbest")=2 then
        response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>该帖子内容不妥已被管理员单帖屏蔽</font><br>================================</DIV>"
        else
        if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
        response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>您没有浏览该精华帖子的权限</font><br>================================</DIV>"
        else
   '======================== 这是UBB代码要加入的 =========================
        session("dateandtime")=rs(4)
        session("UserName")=UserName
        session("id")=AnnounceID
        dim rsr,sqlr,checkreplyname
        set rsr=server.createobject("adodb.recordset")
        sqlr="select TOP 1 AnnounceID from "&TotalUseTable&" where boardid="&boardid&" and rootid="&rootid&" and username='"&cstr(membername)&"'"
        rsr.open sqlr,conn,0,1
        if not rsr.eof then
        session("checkreplyname")=true
        else
        session("checkreplyname")=false
        end if
        rsr.close
        set rsr=nothing
   ' ======================== 到这里结束 =========================
        response.write "<img src="&Forum_info(8)&rs(6)&" border=0 alt=发帖心情> <b>"&htmlencode(rs(3)) &"</b><br><div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">"&dvbcode(rs(5),reUserGroupID,1)&"</div>"
        end if
        end if
 response.write "</td><td bgColor=#fffff1 width=""18%"" valign=""top""><script src=""inc/guajian.js""></script></td></tr></table>"
 response.write "</td>"
 response.write "<td background=""images/haha/center_r.gif"" width=""16""></td>"
'帖子加挂件结束
	'response.write "<td background=""images/haha/center_l.gif"" width=""14"" ></td>"
	'response.write "<td style=""font-size:"&Board_setting(28)&"pt;line-height:"&Board_setting(29)&"pt;WORD-WRAP: break-word;"" bgColor=#fffff1 name=""bbscontent"">"
	'if LockUser=2 then
		'response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>该用户发言已被管理员屏蔽</font><br>================================</DIV>"
	'elseif rs("isbest")=2 then
		'response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>该帖子内容不妥已被管理员单帖屏蔽</font><br>================================</DIV>"
	'else
		'if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
			'response.write "<div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">================================<br><font color=red>您没有浏览该精华帖子的权限</font><br>================================</DIV>"
		'else
			'======================== 这是UBB代码要加入的 =========================
			'session("dateandtime")=rs(4)
			'session("UserName")=UserName
			'session("id")=AnnounceID
			'dim rsr,sqlr,checkreplyname
			'set rsr=server.createobject("adodb.recordset")
			'sqlr="select TOP 1 AnnounceID from "&TotalUseTable&" where boardid="&boardid&" and rootid="&rootid&" and username='"&cstr(membername)&"'"
			'rsr.open sqlr,conn,0,1
			'if not rsr.eof then
				'session("checkreplyname")=true
	 		'else
				'session("checkreplyname")=false
		 	'end if
		  'rsr.close
			'set rsr=nothing
			' ======================== 到这里结束 =========================
			'response.write "<img src="&Forum_info(8)&rs(6)&" border=0 alt=发帖心情>&nbsp;<b>"&htmlencode(rs(3)) &"</b><br><div id=""fontsize_"&rs(0)&""" style=""WORD-WRAP: break-word;"">"&dvbcode(rs(5),reUserGroupID,1)&"</div>"
		'end if
	'end if
	'response.write "</td>"
	'response.write "<td background=""images/haha/center_r.gif"" width=""16""></td>"


	response.write "<td width=20></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td width=20></td>"
	response.write "<td background=""images/haha/center_l.gif"" width=""14"" ></td>"
	response.write "<td bgColor=#fffff1 height=20></td>"
	response.write "<td background=""images/haha/center_r.gif"" width=""16""></td>"
	response.write "<td width=20></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td width=20></td>"
	response.write "<td background=""images/haha/foot_l1.gif"" width=""14"" ></td>"
	response.write "<td background=""images/haha/foot_c.gif"" ><img src=""images/haha/foot_l3.gif""></td>"
	response.write "<td background=""images/haha/foot_r.gif"" width=""16""></td>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td width=20></td>"
	response.write "<td width=14>&nbsp;</td>"
	response.write "<td>"&isOnline(PostUserid,sex)&"  <font color="""&Forum_Body(27)&"""><b>"&username&"</b> "
	if isnull(rs("speak")) or trim(rs("speak"))="" then
		response.write "说道"
	else
		response.write trim(rs("speak"))
	end if
	response.write "</font></td>"
	response.write "<td width=16>&nbsp;</td>"
	response.write "<td width=20></td>"
	response.write "</tr>"
	response.write "</table></td></tr>"
	response.write "<tr>"
	response.write "<td colspan=3 valign=bottom><blockquote>"
	if signflag=1 and lockuser=0 and Cint(forum_setting(42))=1 then
		if sign<>"" then response.write "<p><img src=pic/qianming.gif alt=""这是"&UserName&"的个性签名！请勿挪用！""><br>"&dvsigncode(sign,reUserGroupID)&"</p>"
	end if
if istopic then
if Dispbbs_plus_ads<>"" then
response.write "<hr size=1 color=#739ace>" & "<br>"
response.write Dispbbs_plus_ads
else
Dispbbs_plus_ads="<a href=""http://www.51eline.com"" target=""_blank""><img src=""tiejian.jpg"" alt=""全力打造精彩江湖与论坛！"" border=""0""></a>"
response.write "<hr size=1 color=#739ace>" & "<br>"
response.write Dispbbs_plus_ads
end if
end if
	response.write "</blockquote></td>"
	response.write "</tr></table>"
	response.write "</td></tr>"
	if Cint(GroupSetting(30))=0 then
		userip="*.*.*.*"
	else
		userip=rs("ip")
	end if
	response.write "<tr><td class="&bgcolor&" valign=middle align=center width=175><a href=look_ip.asp?boardid="&boardid&"&userid="&PostUserID&"&ip="&userip&"&action=lookip target=_blank><img align=absmiddle border=0 width=13 height=15 src="""&Forum_info(7)&Forum_TopicPic(20) & """ alt=""点击查看用户来源及管理<br>发帖IP："&userip&"""></a> "&rs(4)&"</td>"
	response.write "<td class="&bgcolor&" valign=middle width=*>"
	response.write "<table width=100% cellpadding=0 cellspacing=0><tr><td align=left valign=middle>&nbsp;"
	if cint(forum_setting(27))=1 then
		%><marquee height=16 scrollamount=2 onmouseover='this.stop()' onmouseout='this.start()' onload="javascript:this.width=screen.width"><%=NewBBS%></marquee><%
	else
		%><SCRIPT language=javascript src=text.js></SCRIPT><%
	end if
	response.write "</td><td align=right nowarp valign=bottom width=120>"
	if founduser then
		if Cint(GroupSetting(10))=1 or Cint(GroupSetting(23))=1 then
			response.write "<a href=editannounce.asp?BoardID="&BoardID&"&replyID="&rs(0)&"&id="&AnnounceID&"&star="&star&"><img src="&Forum_info(7)&"Edit1.GIF border=0 alt=编辑这个帖子 align=absmiddle></a> "
		end if
	end if
	if rs("isbest")=0 then
		besttitle="精华"
		pbtitle="单帖屏蔽"
	elseif rs("isbest")=1 then
		bestTitle="解除精华"
		pbtitle="单帖屏蔽"
	else
		bestTitle="精华"
		pbtitle="解除单帖屏蔽"
	end if
	if Cint(GroupSetting(24))=1 then
		response.write "<a href=""admin_postings.asp?action="&pbTitle&"&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&"""><img src=pic/pb.gif border=0 alt="""&pbTitle&"""></a> "
	end if
	if not istopic then
		if Cint(GroupSetting(11))=1 or Cint(GroupSetting(18))=1 then
			response.write "<a href=""admin_postings.asp?action=删除跟帖&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&"&userid="&postuserid&"""><img src="&Forum_info(7)&Forum_TopicPic(17)&" border=0 alt=注意：本操作将删除单个帖子></a>&nbsp;"
		end if
	else
		if Cint(GroupSetting(12))=1 or Cint(GroupSetting(19))=1 then
			response.write "<a href=""admin_postings.asp?action=复制&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&"""><img src="""&Forum_info(7)&Forum_TopicPic(18)&""" border=0 alt=""复制单个帖子到别的版面""></a>&nbsp;"
		end if
		if Cint(GroupSetting(24))=1 then
			response.write "<a href=""admin_postings.asp?action="&bestTitle&"&BoardID="&BoardID&"&ID="&AnnounceID&"&replyID="&rs(0)&"""><img src="""&Forum_info(7)&Forum_TopicPic(19)&""" border=0 alt="""&bestTitle&"""></a>"
		end if
	end if
	response.write "</td>"
	response.write "<td align=right valign=bottom width=4> </td>"
	response.write "</tr></table>"
	response.write "</td></tr></table>"
page_count = page_count + 1
isagree=""
istopic=false
rs.movenext
loop
else
	ErrMsg=ErrMsg+"<br>"+"<li>您指定的帖子不存在</li>"
	exit sub
end if

response.write "<BR><table cellpadding=0 cellspacing=3 border=0 width="&Forum_body(12)&" align=center><tr><td valign=middle nowrap>本主题帖数<b>"&TopicCount&"</b>，分页："
if skin=0 then
	call DispPageNum(star,PCount,"""?BoardID="&BoardID&"&id="&request("ID")&"&replyID="&replyID&"&star=","&skin="&request("skin")&"""")
else
	response.write "无"
end if
rs.close
set rs=nothing

response.write "</td><td valign=middle nowrap align=right><select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}""><option selected>跳转论坛至...</option>"
myCache.name="BoardJumpList"
if myCache.valid then
response.write myCache.value
else
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
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
response.write "</select></td></tr></table>"

if cint(skin)=1 then
%>
<br>
<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<tr><th align=left width=90% valign=middle> &nbsp;*树形目录</th>
<th width=10% align=right valign=middle height=24 id=tabletitlelink> <a href=#top><img src=pic/gotop.gif border=0>顶端</a>&nbsp;</th></tr>
<%
dim outtext,bytestr,fontcolor
set rs=server.createobject("adodb.recordset")
sql="select AnnounceID,parentID,BoardID,UserName,PostUserid,Topic,DateAndTime,length,RootID,layer,orders,Expression,body,isbest from "&TotalUseTable&" where BoardID="&cstr(BoardID)&" and RootID="&Announceid&" and Locktopic<2 order by RootID desc,orders"
	rs.open sql,conn,1,1
	if not(rs.eof and rs.bof) then
	do while not rs.eof
		if bgcolor="tablebody1" then
			bgcolor="tablebody2"
		else
			bgcolor="tablebody1"
		end if         
		if clng(replyid)=rs(0) then
			fontcolor="<font color="&Forum_body(8)&">"
		else
			fontcolor=""
		end if
		bytestr="("+cstr(rs("length"))
		if rs("Length")-1=1 then
			bytestr=bytestr+"字)"
		else
			bytestr=bytestr+"字)"
		end if
		if rs("length")=0 then
			bytestr="(空)"
		end if
		if rs("layer")>1 then
			for i=2 to rs("layer")
			outtext=outtext & " &nbsp; &nbsp; "
			next
			outtext=outtext & "回复："
		else
			outtext=outtext & "主题："
		end if
		outtext=outtext & "&nbsp;<img src="&Forum_info(8)&rs("Expression")&" height=16 width16>  "
		outtext=outtext &  "<a href='dispbbs.asp?BoardID="&BoardID&"&ID="&cstr(rs("RootID"))&"&replyID="&Cstr(rs("AnnounceID"))&"&skin=1'>" & fontcolor
		if rs("topic")="" or isnull(rs("topic")) then
			if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
				outtext=outtext & "（精华帖，禁止预览）"
			elseif rs("isbest")=2 then
				outtext=outtext & "（单帖屏蔽，禁止预览）"
			else
				if len(rs("body"))>35 then
				outtext=outtext & left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),35)+"..."
				else
				outtext=outtext & reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
				end if
			end if
		else
			if len(rs("Topic"))>35 then
			outtext=outtext & left(htmlencode(rs("Topic")),35)+"..."
			else
			outtext=outtext & htmlencode(rs("Topic"))
			end if
		end if
		response.write "<TR><td class="&bgcolor&" width=""100%"" height=22 colspan=2>"&outtext&"</a><I><font color=gray>"&bytestr&" － <a href=dispuser.asp?name="&htmlencode(rs("UserName"))&" target=_blank title='作者资料'><font color=gray>"&HTMLEncode(rs("UserName"))&"</font></a>，" & Formatdatetime(rs("DateAndTime"),1) & "</font></I></td></tr>"
		rs.movenext
		outtext=""
		loop
	end if
	rs.close
	set rs=nothing
response.write "</table><br>"
end if

if canreply then
%>
<table cellpadding=1 cellspacing=1 class=tableborder1 align=center>
<tr>
<th align=left colspan=2 width=100% valign=middle height=25> &nbsp;*快速回复：<%=htmlencode(topic_1)%></th>
</tr>
<script src="inc/ubbcode.js"></script>
<form action="SaveReAnnounce.asp?method=fastreply&BoardID=<%=BoardID%>" method=POST  name=frmAnnounce onSubmit=submitonce(this)>
<input type=hidden name="followup" value="<%=AnnounceID%>">
<input type=hidden name="RootID" value="<%=AnnounceID%>">
<input type=hidden name="star" value="<%=star%>">
<input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">
<TR><TD noWrap width=175 class=tablebody1><b>你的用户名:</b></TD>
<TD height=30 class=tablebody1><INPUT maxLength=25 size=23 value="<%=membername%>" name=UserName>
&nbsp;&nbsp; <A href="reg.asp">还没注册?</A>  &nbsp;&nbsp;  <b>密码:</b>
<INPUT type=password maxLength=20 size=23 value="<%=memberword%>" name=passwd>
&nbsp;&nbsp; <A href="lostpass.asp">忘记密码?</A></TD></TR>

<TR><td noWrap width=175 class=tablebody2><li><A href=reg.asp>[ 还没注册 ]</A><br><li><A href=lostpass.asp>[ 忘记密码 ]</A><li>[<a href=javascript:openScript('smiles.asp',500,600)><font color=red>加入心情图标</font></a>]<br><li>将放在帖子的前面<br></td>
<td class=tablebody2>
<%for i=0 to Forum_PostFaceNum-1%>
<input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;
<%if i>0 and ((i+1) mod 9=0) then response.write "<br>"%>
<%next%>
</td></tr>
<TR> <TD vAlign=top noWrap class=tablebody1><b>内容</b><br><!--#include file="z_ContentFlag.asp"--><%if board_setting(44)=1 then%>&nbsp;<input type=checkbox name=ArticleRandom value=yes <%if board_setting(45)=1 then%>checked<%end if%>>发帖机遇<%else%><input type=hidden name=ArticleRandom value=no><%end if%></TD>
<TD class=tablebody1><%
if Cint(Board_Setting(4))=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<TEXTAREA name=Content cols=103 rows=8 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交帖子  onkeydown=ctlent()></TEXTAREA>
<!--#include file="z_speak.asp"--></TD></TR>
<TR>
<td class=tablebody2 valign=top colspan=2 style="table-layout:fixed; word-break:break-all"><b>点击表情图即可在帖子中加入相应的表情</B><br>&nbsp;<%
	dim ii
	for i=0 to 14
		if len(i)=1 then ii="0" & i else ii=i
		response.write "<img src="""&Forum_info(10)&Forum_emot(i)&""" border=0 onclick=""insertsmilie('[em"&ii&"]')"" style=""CURSOR: hand"">"
	next
%></td>
<TR>
<TD noWrap class=tablebody1 height=30><INPUT type=checkbox value=yes name=signflag checked>签名<%if board_setting(46)=1 then%> <input type=checkbox name=msgflag value=yes <%if board_setting(47)=1 then%>checked<%end if%> alt="<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
	if isnull(myvip) or myvip<>1 then
		response.write forum_user(28)
	else
		response.write forum_user(29)
	end if
	%>元，您现在有现金<%=mymoney%>元</font>">短信<%else%><input type=hidden name=msgflag value=no><%end if%> <INPUT type=checkbox value=yes name=emailflag alt="<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
	if isnull(myvip) or myvip<>1 then
		response.write forum_user(30)
	else
		response.write forum_user(31)
	end if
%>元，您现在有现金<%=mymoney%>元</font>">邮件</TD>
<TD width="100%" class=tablebody1> 
<input type=Submit value="OK!发表我的回应帖子" name=Submit onclick="if(checkbbs(false)){return true;} return false;">
&nbsp;<input type=button value="预 览" name=Button onclick=gopreview()>&nbsp;<input type=reset name=Clear value=清空内容！>[Ctrl+Enter直接提交帖子] 
</TD>
</TR></FORM>
</TABLE>
<%
end if
dim istoptitle_a
if istopic or skin=0 then
	if istop=2 then istoptitle_a="解除总固顶" else istoptitle_a="总固顶"
	if istop=1 then istoptitle="解固" else istoptitle="固顶"
%><BR>
<TABLE cellSpacing=0 cellPadding=0 width="<%=Forum_body(12)%>" border=0 align=center>
<tr valign=center> <td width =100% align=right> <b>管理选项</b>：
<a href="admin_postings.asp?action=锁定&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=锁定本主题>锁定</a> 
  | <a href="admin_postings.asp?action=解锁&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题解开锁定>解锁</a>
  | <a href="admin_postings.asp?action=提升&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题提升到帖子列表最前面>提升</a>
  | <a href="admin_postings.asp?action=下沉&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题下沉到七天前>下沉</a>
  | <a href="admin_postings.asp?action=删除主题&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=注意：本操作将删除本主题所有帖子，不能恢复>删除</a>
  | <a href="admin_postings.asp?action=移动&BoardID=<%=BoardID%>&ID=<%=Announceid%>&replyID=<%=replyID%>" title=移动主题>移动</a>  |  <a href="admin_postings.asp?action=<%=istopTitle%>&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题<%=istopTitle%>><%=istopTitle%></a>  |  <a href="admin_postings.asp?action=<%=istopTitle_a%>&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=将本主题<%=istopTitle_a%>><%=istopTitle_a%></a>
  | <a href="admin_postings.asp?action=奖励&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=对发起主题用户奖励>奖励</a>
  | <a href="admin_postings.asp?action=惩罚&BoardID=<%=BoardID%>&ID=<%=Announceid%>" title=对发起主题用户惩罚>惩罚</a>
  | <a href="announcements.asp?BoardID=<%=BoardID%>&action=AddAnn">发布公告</a>
</td></tr></table>
<%end if%>
<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
<input type=hidden name=title value=""><input type=hidden name=body value="">
</form>
<script>
function gopreview()
{
<%if voteyes=0 then%>
document.forms[1].body.value=document.forms[0].Content.value;
var popupWin = window.open('preview.asp?boardid=<%=boardid%>', 'preview_page', 'scrollbars=yes,width=750,height=450');
document.forms[1].submit()
<%else%>
document.forms[2].body.value=document.forms[1].Content.value;
var popupWin = window.open('preview.asp?boardid=<%=boardid%>', 'preview_page', 'scrollbars=yes,width=750,height=450');
document.forms[2].submit()
<%end if%>
}
</script>
<script language=javascript>
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.frmAnnounce.submit();}}
}

function submitonce(theform){
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
<%call activeonline()
end sub

sub readRe()
	dim rs1,ID,i,reAnnSplit,reAnnStr
	set rs1=conn.execute("select reAnn from [user] where UserID="&UserID&" and reAnn is not null")
	if not (rs1.eof and rs1.bof) then
		reAnnSplit=split(rs1("reAnn"),"|")
		reAnnStr=""
		for i=0 to ubound(reAnnSplit) step 2
			ID=reAnnSplit(i+1)
			if clng(ID)<>clng(AnnounceID) then
				if reAnnStr<>"" then
					reAnnStr=reAnnStr & "|"
				end if
				reAnnStr=reAnnStr & reAnnSplit(i) & "|" & reAnnSplit(i+1)
			end if
		next
		if reAnnStr="" then
			conn.execute ("update [user] set reAnn=null where UserID="&UserID)
		else
			conn.execute ("update [user] set reAnn='"&reAnnStr&"' where UserID="&UserID)
		end if	
	end if
	rs1.close
	set rs1=nothing
end sub

function isOnline(Userid,sex)
	dim ors
	set ors=conn.execute("select username from online where userid="&userid&" "&userhiddensql&"")
	if ors.eof and ors.bof then
		if sex=1 then
		isonline="<img src=pic/ofMale.gif alt=帅哥哟，离线，有人找我吗？>"
		else
		isonline="<img src=pic/ofFeMale.gif alt=美女呀，离线，快来找我吧！>"
		end if
	else
		if sex=1 then
		isonline="<img src=pic/Male.gif alt=帅哥哟，在线，有人找我吗？>"
		else
		isonline="<img src=pic/FeMale.gif alt=美女呀，在线，快来找我吧！>"
		end if
	end if
	ors.close
	set ors=nothing
end function

Function CurNewTopic(tlen,n,info)
	dim rs,sql
	dim topic
	dim board
	dim NewTopic
	NewTopic=""
	set rs=conn.execute("select top "&n&" username,topic,rootid,boardid,dateandtime,announceid,body,isbest from "&NowUseBbs&" where locktopic<2 ORDER BY announceid desc")
	do while Not RS.Eof
		if renzhen(rs("boardid"),membername)=false then
			topic="<font color=gray>（认证或隐藏论坛帖子，只有认证用户或版主才能查看）</font>"
		elseif VipBoard(rs("boardid"),membername)=false then
			topic="<font color=gray>（VIP或特殊论坛帖子，只有VIP或特殊会员才能查看）</font>"
		else
			if rs("topic")="" or isnull(rs("topic")) then
				if rs("isbest")=1 and Cint(GroupSetting(41))=0 then
					topic="（精华帖，禁止预览）"
				elseif rs("isbest")=2 then
					topic="（单帖屏蔽，禁止预览）"
				else
					if len(rs("body"))>tlen then
						topic=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),tlen)+"..."
					else
						topic=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
					end if
				end if
			else
				if len(rs("Topic"))>tlen then
					topic=left(htmlencode(rs("Topic")),tlen)+"..."
				else
					topic=htmlencode(rs("Topic"))
				end if
			end if
		end if
		if trim(topic)="" then topic="..."
		NewTopic=NewTopic&" <img border=0 src=pic/nologin.gif> <span style=""font-size:9pt;line-height: 15pt""><a href=dispbbs.asp?boardid="&rs(3)&"&ID="&rs(2)&"&replyID="&rs(5)&"&skin=1 target=_blank><FONT color=#0000FF title="&htmlencode(rs(1))&">"&topic&"</FONT></a>"
		select case info
		case 1
			NewTopic=NewTopic&" (<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a> <font color=green>"&month(rs("dateandtime"))&"-"&day(rs("dateandtime"))&" "&formatdatetime(rs("dateandtime"),4)&"</font>)"
		case 2
			NewTopic=NewTopic&" (<font color=green>"&month(rs("dateandtime"))&"-"&day(rs("dateandtime"))&" "&formatdatetime(rs("dateandtime"),4)&"</font>)"
		case 3
			NewTopic=NewTopic&" (<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>)"
		end select
		NewTopic=NewTopic&"</span>"
		RS.MoveNext
	Loop
	rs.close
	set rs=nothing
	CurNewTopic=NewTopic
End Function
%>