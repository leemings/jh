<!--#include file=char.asp-->
<%
'论坛Cookies名称，一般不需要修改，当您的网站启用两个以上的论坛时必须将这些论坛系统的Cookies名称修改为不同
Dim Forum_sn
Dim Mycache
set myCache=new Cache
Forum_sn="aspsky"
if Request.Cookies("iscookies")="" then
	Response.Cookies("iscookies")="0"
	Response.Cookies("iscookies").Expires=date+10
	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3><div align=center><img src=../eline_bg.gif></div>"
	response.end
end if

Dim UserAgent,Stats,ScriptName
UserAgent=Trim(lcase(Request.Servervariables("HTTP_USER_AGENT")))
ScriptName=lcase(request.ServerVariables("PATH_INFO"))
If Instr(lcase(trim(UserAgent)),"teleport")>0 or Instr(lcase(trim(UserAgent)),"webzip")>0 or Instr(lcase(trim(UserAgent)),"flashget")>0 or Instr(lcase(trim(UserAgent)),"offline")>0 Then
	response.redirect "error.htm"
	response.end
end if

Dim copyright,Version,Cookiepath,Badwords,Splitwords,StopReadme,Maxonline,MaxonlineDate
Dim LastTopicNum,LastPostTime,LastPostInfo,lastbbsnum
Dim guestlist
Dim boardtype,lockboard,master_2,todaynum,master_1,boardmasterlist
Dim skinid,skin,tid,tempid
Dim templateused
Dim Forum_info,Forum_setting,Forum_ubb,Forum_body,Forum_ads,Forum_user,Board_Setting
Dim Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_upload
Dim MaxOldID,NowUseBbs,AllPostTable,AllPostTableName,BoardReadme
Dim sqlstr,sqlgroup
Dim mastermsg,Dispbbs_plus_ads
Dim Forum_getre,Forum_smsmsg
Dim backpage,navinfo
Dim Forum_header
Dim membername,memberword,memberclass,userhidden
Dim sql,rs,ors
Dim BoardID,FoundBoard,i
Dim Founderr,Errmsg,sucmsg
Dim Founduser,userid
Dim FoundStyle
Dim UserPermission,FoundUserPer
Dim BoardParentStr,BoardParentID,BoardDepth,BoardChild
Dim Forum_OnlineNum
Dim Forum_userface,Forum_PostFace,Forum_emot
Dim Forum_userfaceNum,Forum_PostFaceNum,Forum_emotNum
Dim OpenTime,isaudit
Forum_OnlineNum=allonline()
FoundUserPer=false
FoundBoard=false
Founduser=false
Founderr=false
FoundStyle=false
if request("BoardID")="" or (not isInteger(request("BoardID"))) or request("boardid")="0" or instr(lcase(trim(scriptname)),"index.asp")>0 then
	BoardID=0
	FoundBoard=false
else
	BoardID=Clng(Request("BoardID"))
	FoundBoard=true
end if
membername=checkStr(request.cookies("aspsky")("username"))
memberword=checkStr(request.cookies("aspsky")("password"))
memberclass=checkStr(request.cookies("aspsky")("userclass"))
userhidden=request.cookies("aspsky")("userhidden")
userid=request.cookies("aspsky")("userid")
if not isnumeric(userhidden) or userhidden="" then userhidden=2
'论坛基本变量调用
if not FoundBoard then
	skinid=request.cookies("skin")("skinid_0")
	if isnumeric(skinid) and not isnull(skinid) and skinid<>"" then
		sqlstr=" id="&skinid&" "
	else
		if not isnumeric(request("skinid")) or request("skinid")="" or request("skinid")="0" then
			sqlstr=" active=1"
		else
			sqlstr=" id="&request("skinid")&" "
		end if
	end if
	sql = "select * from config where "&sqlstr&""
else
	skinid=request.cookies("skin")("skinid_"&boardid)
	if isnumeric(skinid) and not isnull(skinid) and skinid<>"" then
	sql = "select b.boardtype,b.boardmaster,b.todaynum,b.lasttopicnum,b.lastbbsnum,b.LastPost,b.readme,b.Board_Setting,b.ParentID,b.Depth,b.ParentStr,b.child,c.* from board b,config c where c.id="&skinid&" and  b.boardid="&boardid
	else
	sql = "select b.boardtype,b.boardmaster,b.todaynum,b.lasttopicnum,b.lastbbsnum,b.LastPost,b.readme,b.Board_Setting,b.ParentID,b.Depth,b.ParentStr,b.child,c.* from board b inner join config c on b.sid=c.id where b.boardid="&boardid
	end if
end if
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	response.write "<meta HTTP-EQUIV=REFRESH CONTENT='1; URL=Ecookies.asp'><font style=font-size:9pt color=#6699FF>错误的系统参数，很可能您指定了错误的论坛名称，请从有效连接进入。<br>或者您选择了已经被删除的自选风格，请<a href=Ecookies.asp>点击清空您在本论坛的Cookies信息</a>或等待一秒后自动清空。</font>"
	response.end
else
	Forum_info=split(rs("Forum_info"),",")
	dim uForum_info
	uForum_info=ubound(Forum_info)
	if uForum_info<12 then
		redim preserve Forum_info(12)
		Forum_info(12)=now()
	end if
	Forum_Setting=split(rs("Forum_setting"),",")
	if Forum_Setting(12)="" then Forum_Setting(12)="0"
	if Forum_Setting(13)="" then Forum_Setting(13)="0"
	if Forum_Setting(16)="" then Forum_Setting(16)="0"
	if Forum_Setting(17)="" then Forum_Setting(17)="1"
	if Forum_Setting(18)="" then Forum_Setting(18)="1"
	if Forum_Setting(27)="" then Forum_Setting(27)="0"
	if Forum_Setting(39)="" then Forum_Setting(39)=Forum_Setting(38)
	if Forum_Setting(43)="" then Forum_Setting(43)="0"
	if Forum_Setting(45)="" then Forum_Setting(45)="0"
	if Forum_Setting(58)="" then Forum_Setting(58)=Forum_Setting(57)
	if Forum_Setting(71)="" then Forum_Setting(71)="0"
	if Forum_Setting(72)="" then Forum_Setting(72)="00"
	dim uForum_Setting
	uForum_Setting=ubound(Forum_Setting)
	if uForum_Setting<77 then
		redim preserve Forum_Setting(77)
		Forum_Setting(73)="0"
		Forum_Setting(74)="1"
		Forum_Setting(75)="1"
		Forum_Setting(76)="1"
		Forum_Setting(77)="3"
	end if
	Forum_ads=split(rs("Forum_ads"),"$")
	Forum_body=split(rs("Forum_body"),"|||")
	Forum_user=split(rs("Forum_user"),",")
Dispbbs_plus_ads=rs("dispbbs_ad")
	dim uForum_user
	uForum_user=ubound(Forum_user)
	if uForum_user<39 then
		redim preserve Forum_user(39)
		Forum_user(18)="100"
		Forum_user(19)="50"
		Forum_user(20)="50"
		Forum_user(21)="25"
		Forum_user(22)="20"
		Forum_user(23)="10"
		Forum_user(24)="40"
		Forum_user(25)="20"
		Forum_user(26)="4"
		Forum_user(27)="2"
		Forum_user(28)="40"
		Forum_user(29)="20"
		Forum_user(30)="0"
		Forum_user(31)="0"
		Forum_user(32)="10"
		Forum_user(33)="5"
		Forum_user(34)="2"
		Forum_user(35)="1"
		Forum_user(36)="2"
		Forum_user(37)="1"
		Forum_user(38)="2"
		Forum_user(39)="1"
	end if
	if uForum_user<42 then
		redim preserve Forum_user(42)
		Forum_user(40)="5"
		Forum_user(41)="1"
		Forum_user(42)="1"
	end if

	Forum_userface=split(rs("Forum_userface"),"|")
	Forum_PostFace=split(rs("Forum_PostFace"),"|")
	Forum_emot=split(rs("Forum_emot"),"|")
	Forum_userfaceNum=ubound(Forum_userface)
	Forum_PostFaceNum=ubound(Forum_PostFace)
	Forum_emotNum=ubound(Forum_emot)

	Copyright=rs("Forum_copyright")%>
	<!--#include file=version.asp-->
	<%badwords=rs("badwords")
	Splitwords=rs("Splitwords")
	StopReadme=rs("StopReadme")
	Forum_pic=split(rs("Forum_pic"),",")
	dim uForum_Pic
	uForum_Pic=ubound(Forum_Pic)
	if uForum_Pic<16 then
		redim preserve Forum_Pic(16)
		Forum_Pic(16)="ao2.gif"
	end if
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_upload=rs("Forum_upload")
	Forum_ubb=split(rs("Forum_ubb"),",")
	Forum_sn=rs("Forum_sn")
	MaxOldID=rs("Maxoldid")
	NowUseBbs=rs("NowUseBbs")
	AllPostTable=split(rs("AllPostTable"),"|")
	AllPostTableName=split(rs("AllPostTableName"),"|")

	OpenTime=split(forum_setting(70),"|")
	if Cint(forum_setting(69))=1 and ubound(OpenTime)=1 then
		if IsNumeric(OpenTime(0)) and IsNumeric(OpenTime(1)) then
			if Hour(Now)<Cint(OpenTime(0)) or Hour(Now)>Cint(OpenTime(1)) then
				response.write "本论坛在<B>"&OpenTime(0)&"</B>至<B>"&OpenTime(1)&"</B>点开放，请在规定时间内访问，谢谢。"
				response.end
			end if
		end if
	end if
	if not FoundBoard then
		cookiepath=rs("cookiepath")
		skinid=rs("id")
		tid=rs("tid")
		stats="论坛信息"
		Maxonline=rs("Maxonline")
		MaxonlineDate=rs("MaxonlineDate")
		if Clng(forum_setting(26))>0 then
			if Forum_OnlineNum>Clng(forum_setting(26)) then
				if membername="" then
				response.write "当前论坛在线已经超过<B>"&forum_setting(26)&"</B>人，请稍后访问。"
				response.end
				else
				set ors=conn.execute("select username from online where username='"&membername&"'")
				if rs.eof and rs.bof then
				response.write "当前论坛在线已经超过<B>"&forum_setting(26)&"</B>人，请稍后访问。"
				response.end
				end if
				end if
			end if
		end if
	else
		LastPostInfo=split(rs(5),"$")
		Board_Setting=split(rs("board_setting"),",")
		dim uboard_setting
		uboard_setting=ubound(board_setting)
		if uboard_setting<61 then
			redim preserve board_setting(61)
			for i=uboard_setting to 61
				board_setting(i)=0
			next
		end if
		for i=44 to 61
			if i>=50 and i<=52 then
				if board_setting(i)="" then board_setting(i)=0
			else
				if board_setting(i)="" then board_setting(i)=1
			end if
		next
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		LastPostTime=LastPostInfo(2)
		boardtype=rs("boardtype")
		lockboard=Board_Setting(0)
		boardmasterlist=rs("boardmaster")
		todaynum=rs(2)
		LastTopicNum=rs("LastTopicNum")
		LastBbsNum=rs("LastBbsNum")
		tid=rs("tid")
		stats=boardtype
		BoardReadme=rs("Readme")
		BoardParentStr=rs("ParentStr")
		BoardParentID=rs("ParentID")
		BoardDepth=rs("depth")
		BoardChild=rs("child")
		skinid=rs("id") 
		if datediff("d",LastPostTime,Now())>0 then
			conn.execute("update board set todaynum=0 where boardid="&boardid)
		end if
		isaudit=Cint(Board_Setting(3))
		OpenTime=split(Board_Setting(22),"|")
		if Cint(Board_Setting(21))=1 and ubound(OpenTime)=1 then
			if IsNumeric(OpenTime(0)) and IsNumeric(OpenTime(1)) then
				if Hour(Now)<Cint(OpenTime(0)) or Hour(Now)>Cint(OpenTime(1)) then
					response.write "本论坛在<B>"&OpenTime(0)&"</B>至<B>"&OpenTime(1)&"</B>点开放，请在规定时间内访问，谢谢。"
					response.end
				end if
			end if
		end if
		if Clng(Board_Setting(18))>0 then
			if Forum_OnlineNum>Clng(Board_Setting(18)) then
				if membername="" then
				response.write "当前论坛在线已经超过<B>"&Board_Setting(18)&"</B>人，请稍后访问。"
				response.end
				else
				set ors=conn.execute("select username from online where username='"&membername&"'")
				if rs.eof and rs.bof then
				response.write "当前论坛在线已经超过<B>"&Board_Setting(18)&"</B>人，请稍后访问。"
				response.end
				end if
				end if
			end if
		end if
	end if
end if
rs.close
set rs=nothing

if cint(Forum_Setting(19))=1 then
	Dim SplitReflashPage
	Dim DoReflashPage
	DoReflashPage=false
	if trim(forum_setting(64))<>"" then
		SplitReflashPage=split(forum_setting(64),"|")
		for i=0 to ubound(SplitReflashPage)
			if instr(lcase(trim(scriptname)),lcase(trim(SplitReflashPage(i))))>0 then
				DoReflashPage=true
				exit for
			end if
		next
	end if
	if (not isnull(session("ReflashTime"))) and cint(Forum_Setting(20))>0 and DoReflashPage then
		if DateDiff("s",session("ReflashTime"),Now())<cint(Forum_Setting(20)) then
   		response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&Forum_Setting(20)&">本页面起用了防刷新机制，请不要在<b><font color=ff0000>"&Forum_Setting(20)&"</font></b>秒内连续刷新本页面<BR>正在打开页面，请稍候……"
		response.end
		else
		session("ReflashTime")=Now()
		end if
	elseif isnull(session("ReflashTime")) and cint(Forum_Setting(20))>0 and DoReflashPage then
		Session("ReflashTime")=Now()
	end if
end if

if instr(lcase(trim(scriptname)),"admin")=0 and instr(lcase(trim(scriptname)),"login")=0 and instr(lcase(trim(scriptname)),"chklogin")=0 then
	if cint(Forum_Setting(21))=1 then
	Response.write StopReadme
	response.end
	end if

	if LockIP(Request.ServerVariables("REMOTE_ADDR")) then
	response.write "您的IP已经被限制不能访问本论坛，请和管理员联系"&Forum_info(5)&"。"
	response.end
	end if
end if

Server.ScriptTimeOut=Forum_Setting(1)

Rem 用户信息
Dim vipuser,boardmaster,master,superboardmaster
Dim reboardid,reid,recount
Dim LastLogin,myuserep,mysex,myusercp,mymoney,mypower,myArticle,myClass,myaddDate,myusergroup,myvip,myuserscore,LastViewTime,myTodayTopic,myTodayReply,myUserPPD
dim vipdate,userwealth2,userep2,usercp2,userpower2,article2,userscore2
Dim UserGroupID,GroupSetting
dim isvip,vipmsgcontent,rsvip
vipuser=false
boardmaster=false
superboardmaster=false
master=false
myClass=0
myuserep=0
myusercp=0
mypower=0
myArticle=0
mymoney=0
myaddDAte=Now()
myusergroup=""
myvip=0
mysex=1
myuserscore=0
myTodayTopic=0
myTodayReply=0
myUserPPD=0
if userid<>"" and isInteger(userid) then
	sql="select u.userclass,u.userid,u.userpassword,u.lastlogin,u.showRe,u.reAnn,u.userGroupID,g.GroupSetting,u.username,u.userWealth,u.Userep,u.Usercp,u.sex,u.UserPower,u.Article,u.addDAte,u.usergroup,u.vip,u.vipdate,u.userwealth2,u.userep2,u.usercp2,u.userpower2,u.article2,u.userscore,u.userscore2,u.LastViewTime,u.TodayTopic,u.TodayReply,u.UserPPD,u.UserLastIP from [user] u inner join UserGroups G on u.UserGroupID=g.UserGroupID where u.userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founduser=false
	else

'shinzeal使用UserTrueIP使结果更准确
dim UserTrueIP 
UserTrueIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserTrueIP = "" Then UserTrueIP = Request.ServerVariables("REMOTE_ADDR")
UserTrueIP = checkStr(UserTrueIP)
if Forum_setting(72)="" then Forum_setting(72)=1 'shinzeal加入Forum_setting(72)作为cookies欺骗开关 
if rs(15)<>UserTrueIP and ((Clng(Forum_setting(72))=1 and rs(6)<4) or Clng(Forum_setting(72))=2) then 'shinzeal判断两次IP的变化，防止cookies欺骗 
Founduser=False
else


		if trim(rs(2))=trim(memberword) and lcase(trim(membername))=lcase(trim(rs(8))) then
			founduser=true
			select case rs(6)
			case 8
				vipuser=true
			case 3
				boardmaster=true
			case 2
				superboardmaster=true
			case 1
				master=true
			end select
			userid=rs(1)
			lastlogin=rs(3)
			LastViewTime=rs("LastViewTime")
			if (rs(4)=1 or rs(4)=true) and rs(5)<>"" then
				reBoardID=split(rs(5),"|")(0)
				reID=split(rs(5),"|")(1)
				reCount=(ubound(split(rs(5),"|"))+1)\2
			end if
			UserGroupID=rs(6)
			myuserep=rs(10)
			mysex=rs(12)
			myusercp=rs(11)
			mymoney=rs(9)
			mypower=rs(13)
			myArticle=rs(14)
			myClass=rs(0)
			myaddDate=rs(15)
			myusergroup=rs(16)
			myvip=rs(17)
			if isnull(myvip) then
				myvip=0
			end if
			myuserscore=rs("userscore")
			vipdate=rs("vipdate")
			userwealth2=rs("userwealth2")
			userep2=rs("userep2")
			usercp2=rs("usercp2")
			userpower2=rs("userpower2")
			myTodayTopic=rs("TodayTopic")
			myTodayReply=rs("TodayReply")
			myUserPPD=rs("UserPPD")
			userscore2=rs("userscore2")
			if myvip=1 or myvip=3 then
				set rsvip=conn.execute("select vipmod,vipdate,wealth2,userep2,usercp2,power2,article2,userscore2,notifydays from [Vip]")
				isvip=cint(rsvip("vipdate")-datediff("d",rs("vipdate"),now))
				if isvip<=0 then
					if rsvip("vipmod")<>1 then
						if (mymoney-userwealth2>=rsvip("wealth2") or rsvip("wealth2")<=0) and (myuserep-userep2>=rsvip("userep2") or rsvip("userep2")<=0) and (myusercp-usercp2>=rsvip("usercp2") or rsvip("usercp2")<=0) and (mypower-userpower2>=rsvip("power2") or rsvip("power2")<=0) and (myarticle-article2>=rsvip("article2") or rsvip("article2")<=0) and (myuserscore-userscore2>=rsvip("userscore2") or rsvip("userscore2")<=0) then
							conn.execute ("update [user] set vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article,userscore2=userscore where userid="&userid)
							vipmsgcontent="您的VIP资格到期，系统已经将您自动加为本期VIP会员，谢谢您对"& forum_info(0) &"的支持！"
							conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&membername&"','"&forum_info(0)&"','您的VIP会员资格续期成功！','"&vipmsgcontent&"',Now(),0,1)")
						else
							conn.execute ("update [User] set vip=0 where  userid="&userid)
							vipmsgcontent="您的VIP资格已经过期，您现在尚未满足再次申请的条件，谢谢您对"& forum_info(0) &"的支持！"&CHR(10)&"如果您愿意，可以在满足条件后重新申请VIP会员！"
							conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&membername&"','"&forum_info(0)&"','您的VIP会员资格续期失败！','"&vipmsgcontent&"',Now(),0,1)")
							myvip=0
						end if
					else
						conn.execute ("update [User] set vip=0,vipdate=now() where  userid="&userid)
						vipmsgcontent="您的VIP资格已经过期，谢谢您对"& forum_info(0) &"的支持！"&CHR(10)&"如果您愿意，您可以再次申请VIP会员！"
						conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&membername&"','"&forum_info(0)&"','您的VIP会员资格已经过期！','"&vipmsgcontent&"',Now(),0,1)")
						myvip=0
					end if
				elseif cint(rsvip("vipdate")-datediff("d",rs("vipdate"),now))<=rsvip("notifydays") then
					if request.cookies("vip_"&userid)<>"1" then
						Response.Cookies("vip_"&userid)="1"
						Response.Cookies("vip_"&userid).Expires=date+1
						if (mymoney-userwealth2>=rsvip("wealth2") or rsvip("wealth2")<=0) and (myuserep-userep2>=rsvip("userep2") or rsvip("userep2")<=0) and (myusercp-usercp2>=rsvip("usercp2") or rsvip("usercp2")<=0) and (mypower-userpower2>=rsvip("power2") or rsvip("power2")<=0) and (myarticle-article2>=rsvip("article2") or rsvip("article2")<=0) and (myuserscore-userscore2>=rsvip("userscore2") or rsvip("userscore2")<=0) then
							'vipmsgcontent="您的VIP资格还有"&(cint(rsvip("vipdate")-datediff("d",rs("vipdate"),now)))&"天就要到期了，您现在已经具备了再次成为VIP会员的资格，恭喜您，谢谢您对"& forum_info(0) &"的支持！"
							'conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&membername&"','"&forum_info(0)&"','您的VIP会员资格续期成功！','"&vipmsgcontent&"',Now(),0,1)")
						else
							vipmsgcontent="您的VIP资格还有"&(cint(rsvip("vipdate")-datediff("d",rs("vipdate"),now)))&"天就要到期了，您现在还没有具备再次成为VIP会员的资格，加油呀，谢谢您对"& forum_info(0) &"的支持！"
							conn.execute("insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&membername&"','"&forum_info(0)&"','您的VIP会员资格续期失败！','"&vipmsgcontent&"',Now(),0,1)")
						end if
					end if
				end if
				rsvip.close
				set rsvip=nothing
			end if
			GroupSetting=split(rs(7),",")
			if userhidden=2 and DateDiff("s",rs(3),Now())>Clng(Forum_Setting(8))*60 then
				if datediff("d",rs(3),Now())=0 then
					conn.execute("update [user] set UserLastIp='"&replace(Request.ServerVariables("REMOTE_ADDR"),"'","")&"',LastLogin=Now() where userid="&userid)
				else
					if cint(forum_setting(73))>0 then
						conn.execute("update [user] set UserLastIp='"&replace(Request.ServerVariables("REMOTE_ADDR"),"'","")&"',LastLogin=Now(),UserPPD=UserPPD+"&forum_setting(74)&",TodayTopic=0,TodayReply=0 where userid="&userid)
						myTodayTopic=0
						myTodayReply=0
					else
						conn.execute("update [user] set UserLastIp='"&replace(Request.ServerVariables("REMOTE_ADDR"),"'","")&"',LastLogin=Now() where userid="&userid)
					end if
				end if
				LastLogin=Now()
			end if
		else
			founduser=false
		end if
end if
	end if
	rs.close
	set rs=nothing
end if
if not founduser then
	founduser=false
	userid=0
	set rs=conn.execute("select GroupSetting from UserGroups where UserGroupID=7")
	GroupSetting=split(rs(0),",")
	UserGroupID=7
	rs.close
	set rs=nothing
end if
if boardid>=0 then GetBoardPermission
'是否继承上级版主，顺带取出上级论坛版面信息
'最多只取向上的10级版面信息
if FoundBoard and BoardParentID>0 then
	Dim FBoardID(9),FBoardName(9),FBoardParentID(9),FBoardMaster
	FBoardmaster=false
	i=0
	set rs=conn.execute("select boardid,boardtype,boardmaster,ParentID from board where boardid in ("&BoardParentStr&") order by orders")
	do while not rs.eof
		FBoardID(i)=rs(0)
		FBoardName(i)=rs(1)
		FBoardParentID(i)=rs(3)
		if Cint(Board_Setting(40))=1 and not FBoardMaster then
			if instr("|"&lcase(trim(rs(2)))&"|","|"&lcase(trim(membername))&"|")>0 then
				FBoardMaster=true
			else
				FBoardMaster=false
			end if
		end if
		i=i+1
		if i>9 then exit do
	rs.movenext
	loop
	set rs=nothing
	if master then
		boardmaster=true
	else
		if FBoardMaster then
			boardmaster=true
		else
			if instr("|"&lcase(trim(boardmasterlist))&"|","|"&lcase(trim(membername))&"|")>0 then
				boardmaster=true
			else
				boardmaster=false
			end if
		end if
	end if
elseif FoundBoard and FoundUser then
	if master then
		boardmaster=true
	else
		if instr("|"&lcase(trim(boardmasterlist))&"|","|"&lcase(trim(membername))&"|")>0 then
		boardmaster=true
		else
		boardmaster=false
		end if
	end if
end if
%>