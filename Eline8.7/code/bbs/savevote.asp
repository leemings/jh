<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<!--#include file="md5.asp"-->
<%
dim announceid
dim UserName
dim userPassword
dim useremail
dim Topic
dim body
dim dateTimeStr
dim ip
dim Expression
dim signflag
dim mailflag
dim msgflag
dim boardstat
dim usercookies
dim votetype,vote,votenum
dim vote_1,votelen,votenumlen,j
dim votetimeout
dim rootid
dim ihaveupfile,upfileinfo,upfilelen
dim ArticleRandom
dim speak
dim msgNotify

stats="发表投票"

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
IP=replace(Request.ServerVariables("REMOTE_ADDR"),"'","")
Expression=Checkstr(Request.Form("Expression"))
Topic=Checkstr(request.Form("subject"))
Body=Checkstr(request.Form("Content"))
UserName=Checkstr(trim(request.Form("username")))
Signflag=Checkstr(trim(request.Form("signflag")))
mailflag=Checkstr(trim(request.Form("emailflag")))
msgflag=Checkstr(trim(request.Form("msgflag")))
ArticleRandom=Checkstr(trim(request.Form("ArticleRandom")))
speak=Checkstr(trim(request.Form("speak")))
msgNotify=Checkstr(trim(request.Form("msgNotify")))
if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
else
	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
end if

if request.form("upfilerename")<>"" then
	ihaveupfile=1
	upfileinfo=replace(request.form("upfilerename"),"'","")
	upfileinfo=replace(upfileinfo,";","")
	upfileinfo=replace(upfileinfo,"--","")
	upfileinfo=replace(upfileinfo,")","")
	upfilelen=len(upfileinfo)
	upfileinfo=left(upfileinfo,upfilelen-1)
else
	ihaveupfile=0
end if
votetype=Checkstr(request.Form("votetype"))
vote=Checkstr(trim(replace(request.Form("vote"),"|","")))

rem -----检查user输入数据的合法性------	
dim num1,rndnum,k
if signflag="yes" then
	signflag=1
else
	signflag=0
end if
if mailflag="yes" then
	mailflag=1
else
	mailflag=0
end if
if msgflag="yes" then
	msgflag=1
else
	msgflag=0
end if
if cint(Board_Setting(30))=1 then
	if not (isnull(session("lastpost")) or boardmaster or master or superboardmaster) then
		if DateDiff("s",session("lastpost"),Now())<cint(Board_Setting(31)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>本论坛限制发帖距离时间为"&Board_Setting(31)&"秒，请稍后再发。"
   		FoundErr=True
		end if
	end if
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
	FoundErr=True
end if
if UserName="" or UserPassWord="" then
	username=membername
	UserPassWord=memberword
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入姓名(长度不能大于20)"
	FoundErr=True
end if
if Topic="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>主题不应为空。"
elseif strLength(topic)>50 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>主题长度不能超过50"
end if
if strLength(body)>Clng(Board_Setting(16)) then
	ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于" & CSTR(Board_Setting(16)) & "bytes"
	FoundErr=true
end if
if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>没有填写内容。"
	FoundErr=true
end if
if vote="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入投票内容"
	FoundErr=true
else
	vote=split(vote,chr(13)&chr(10))
	j=0
	for i = 0 to ubound(vote)
		if not (vote(i)="" or vote(i)=" ") then
			vote_1=""&vote_1&""&vote(i)&"|"
			j=j+1
		end if
		if i>cint(Board_Setting(32))-2 then exit for
	next
	for k = 1 to j
		votenum=""&votenum&"0|"
	next
	votelen=len(vote_1)
	votenumlen=len(votenum)
	votenum=left(votenum,votenumlen-1)
	vote=left(vote_1,votelen-1)
end if
if not isnumeric(request("votetimeout")) then
	ErrMsg=ErrMsg+"<Br>"+"<li>错误的时间参数。"
	FoundErr=true
else
	if request("votetimeout")="0" then
		votetimeout=dateadd("d",9999,Now())
	else
		votetimeout=dateadd("d",request("votetimeout"),Now())
	end if
	votetimeout=replace(replace(CSTR(votetimeout+Forum_Setting(0)/24),"上午",""),"下午","")
end if
session("lastpost")=Now()

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call main()
	if founderr then 
		call nav()
		call head_var(1,BoardDepth,0,0)
		call dvbbs_error()
	end if
end if
call footer()

sub main()
if (Cint(Board_Setting(43))=1 or Cint(Board_Setting(0))=1) and not (master or superboardmaster or boardmaster) then
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发帖。"
	founderr=true
	exit sub
end if
if Cint(GroupSetting(8))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表投票的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
	exit sub
end if
if cint(Board_Setting(2))=1 then
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

usercookies=request.Cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies=3

if chkuserlogin(username,userpassword,usercookies,2)=false then
	errmsg=errmsg+"<br>"+"<li>您的用户名并不存在，或者您的密码错误，或者您的帐号已被管理员锁定。"
	founderr=true
	exit sub
end if

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
		exit sub
	end if
end if
if not isnull(msgNotify) and msgNotify<>"" then
	set rs=conn.execute("select username from [user] where username='"&msgNotify&"'")
	if rs.bof or rs.eof then
		set rs=nothing
		Errmsg=ErrMsg+"<Br>"+"<li>您指定的通知送达者不存在！"
		founderr=true
		exit sub
	end if
	rs.close
	dim notifyusemoney
	notifyusemoney=0
	if not isnull(msgNotify) and msgNotify<>"" then
		set rs=conn.execute("select username from [user] where username='"&msgNotify&"'")
		if rs.bof or rs.eof then
			set rs=nothing
			Errmsg=ErrMsg+"<Br>"+"<li>您指定的通知送达者不存在！"
			founderr=true
			exit sub
		end if
		rs.close
		if isnull(myvip) or myvip<>1 then
			notifyusemoney=clng(forum_user(32))
		else
			notifyusemoney=clng(forum_user(33))
		end if
		if mymoney<notifyusemoney then
			Errmsg=ErrMsg+"<Br>"+"<li>您的现金不够支付短信引起注意的费用！"
			founderr=true
			exit sub
		end if
	end if
end if
dim mailusemoney,msgusemoney
mailusemoney=0:msgusemoney=0
if msgflag=1 then
	if isnull(myvip) or myvip<>1 then
		msgusemoney=clng(forum_user(28))
	else
		msgusemoney=clng(forum_user(29))
	end if
end if
if mailflag=1 then
	if isnull(myvip) or myvip<>1 then
		mailusemoney=clng(forum_user(30))
	else
		mailusemoney=clng(forum_user(31))
	end if
end if
if msgflag=1 or mailflag=1 then
	if mymoney<msgusemoney+mailusemoney then
		Errmsg=ErrMsg+"<Br>"+"<li>您的现金不够支付短信通知或邮件通知的费用！"
		founderr=true
		exit sub
	end if
end if
select case cint(forum_setting(73))
case 1
	if myUserPPD*cint(forum_setting(74))<=myTodayTopic then
		Errmsg=ErrMsg+"<Br>"+"<li>您今天发表的主题数(<b>"&myTodayTopic&"</b>)已经超过了限制(<b>"&(myUserPPD*cint(forum_setting(74)))&"</b>)，不能继续发表主题！"
		founderr=true
		exit sub
	end if
case 3
	if myUserPPD*cint(forum_setting(74))<=myTodayTopic+myTodayReply then
		Errmsg=ErrMsg+"<Br>"+"<li>您今天发表的帖子数(<b>"&(myTodayTopic+myTodayReply)&"</b>)已经超过了限制(<b>"&(myUserPPD*cint(forum_setting(74)))&"</b>)，不能继续发表帖子！"
		founderr=true
		exit sub
	end if
end select
dim locktopic
if master or superboardmaster or boardmaster then
	isaudit=0
	locktopic=0
elseif isaudit=1 and not (master or superboardmaster or boardmaster) then
	isaudit=1
	locktopic=3
else
	isaudit=0
	locktopic=0
end if
dim LastPost,LastPost_1,uploadpic_n,Forumupload,u
dim LastPostTimes,voteid
DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"上午",""),"下午","")

'插入投票记录
conn.execute("insert into vote (vote,votenum,votetype,timeout) values ('"&vote&"','"&votenum&"',"&votetype&",'"&votetimeout&"')")
set rs=conn.execute("select max(voteid) from vote")
voteid=rs(0)
'插入主题表
sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,isvote,PollID,voteTotal,PostTable,locktopic) values ('"&topic&"',"&boardid&",'"&username&"',"&userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"',1,"&voteid&",0,'"&NowUseBbs&"',"&locktopic&")"
conn.execute(sql)

set rs=conn.execute("select max(topicid) from topic")
rootid=rs(0)
Forum_upload="gif,jpg,jpeg,bmp,zip,rar,html,swf,mid,midi,flash,rm,ra,asf,avi,wmv,exe,xls,dos,ftp,mp3,m3u,txt,mdb,dll,IMG"
Forumupload=split(Forum_upload,",")
for u=0 to ubound(Forumupload)
	if instr(body,"[upload="&Forumupload(u)&"]") or instr(body,"."&Forumupload(u)&"") or instr(body,"["&Forumupload(u)&"]")  then
		uploadpic_n=Forumupload(u)
		exit for
	end if
next
if instr(body,"viewfile.asp?ID=") then uploadpic_n="down"
'插入回复表
Sql="insert into "&NowUseBbs&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,isaudit,speak) values ("&boardid&",0,'"&username&"','"&topic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&rootid&",1,0,'"&ip&"','"&Expression&"',"&locktopic&","&signflag&","&(mailflag+msgflag*2)&",0,"&userid&","&ihaveupfile&","&isaudit&",'"&speak&"')"
conn.execute(sql)

set rs=conn.execute("select max(Announceid) from "&NowUseBbs&"")
Announceid=rs(0)

if ihaveupfile=1 then conn.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Topic&"' where F_ID in ("&upfileinfo&")")
LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(reUBBCode(body,false),20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost=replace(LastPost,"'","")
conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rootid)

LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(reUBBCode(Topic,false),20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost_1=replace(LastPost_1,"'","")
if isaudit=0 then
Dim UpdateBoardID
UpdateBoardID=BoardParentStr & "," & BoardID
if datediff("d",LastPostTime,Now())=0 then
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
else
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
end if
conn.execute(sql)

Dim updateinfo
set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
LastPostTimes=split(rs(0),"$")
LastPostTime=LastPostTimes(2)
if not isdate(LastPostTime) then LastPostTime=Now()
if datediff("d",LastPostTime,Now())=0 then
if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost&"' "&updateinfo&" where active=1"
else
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost&"' where active=1"
end if
conn.execute(sql)
end if
conn.execute("update [user] set userwealth=userwealth-"&(msgusemoney+mailusemoney)&" where username='"&membername&"'")

if lcase(trim(username))<>lcase(trim(msgNotify)) and trim(msgNotify)<>"" and Board_Setting(48)=1 then
	dim rs2,sql2,boardname,msgcontent
	boardname=conn.execute("select boardtype from board where boardid="&boardid)(0)
	msgcontent=msgNotify&"，你好："&chr(10)&"　　我于[b]"&now&"[/b]在[color=red]"&boardname&"[/color]版面发表了一篇题为[color=red]《"&topic&"》[/color]的投票，我希望你可以去看看。"&chr(10)&"[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"][color=blue][b][u]查看帖子内容[/u][/b][/color][/URL][/align]"&chr(10)&"[align=right]"&username&"[/align]"
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from [message]"      
	rs2.open sql2,conn,1,3      
	rs2.addnew      
	rs2("sender")=username
	rs2("incept")=msgNotify
	rs2("title")="请注意"
	rs2("content")=msgcontent
	rs2("flag")=0 
	rs2("issend")=1
	rs2.update
	rs2.close
	set rs2=nothing
	conn.execute("update [user] set userwealth=userwealth-"&notifyusemoney&" where username='"&membername&"'")
end if
response.cookies("LastView")("Topic_"&RootID&"_"&BoardID)=now()

call nav()
call head_var(1,BoardDepth,0,0)
dim PostRetrunName,PostRetrun
select case Board_Setting(17)
case 1
	response.write "<meta http-equiv=refresh content=""5;URL=index.asp"">"
	PostRetrunName="首页"
case 2
	response.write "<meta http-equiv=refresh content=""5;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="您所发布的论坛"
case 3
	if isaudit=1 then
	response.write "<meta http-equiv=refresh content=""5;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="您发布的帖子必须经管理员审核后方可见"
	else
	response.write "<meta http-equiv=refresh content=""5;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&""">"
	PostRetrunName="您所发表的帖子"
	end if
end select
set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">状态：<%=stats%></th>
</tr><tr><td width="100%" class=tablebody1>
本页面将在5秒后自动返回<%=PostRetrunName%>，<b>您可以选择以下操作：</b><br><ul>
<li><a href="index.asp">返回首页</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><%if isaudit=1 then%><%=PostRetrunName%><%else%><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a><%end if%></li>
</ul></td></tr></table>
<%
if ArticleRandom="yes" then
	call eddulf()
end if
end sub
%>
<!--#include file="z_eddul.asp"-->