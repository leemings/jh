<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/email.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<!--#include file="md5.asp"-->
<%sub haveRe()
	dim userid1,rs1,child,i,reAnnSplit,flag
	sql="select postuserid,child from topic where topicID="&rootID
	set rs1=conn.execute (sql)
	userid1=rs1(0)
	child=rs1(1)
	set rs1=nothing
	
	if userid<>userid1 and child>0 then
		sql="select reAnn,showre from [user] where userid="&userid1&" "
		set rs1=conn.execute (sql)
		if not (rs1.eof and rs1.bof) then
			if rs1("showre") then
				if not isnull(rs1("reAnn")) and rs1("reAnn")<>"" then
					reAnnSplit=split(rs1("reAnn"),"|")
					flag=false
					for i=0 to ubound(reAnnSplit) step 2
						if cint(reAnnSplit(i))=cint(boardid) and clng(reAnnSplit(i+1))=clng(rootID) then
							flag=true
							exit for
						end if
					next
					if not flag and len(rs1("reAnn")&"|"&boardID&"|"&rootID)<=50 then
						sql="update [user] set reAnn=reAnn+'"&"|"&boardID&"|"& rootID &"' where userid="&userid1
						conn.execute sql
					end if
				else
					sql="update [user] set reAnn='"&boardID&"|"& rootID &"' where userid="&userid1
					conn.execute sql
				end if
			end if
		end if
		set rs1=nothing
	end if
end sub
%>

<%
dim UserName
dim userPassword
dim useremail
dim Topic,PostTopic
dim body
dim dateTimeStr
dim ParentID
dim RootID
dim iLayer
dim iOrders
dim ip
dim announceid
dim Expression
dim signflag
dim mailflag
dim msgflag
dim Email,mailbody,SendMail
dim boardstat
dim usercookies
dim LastPost,LastPost_1,UpLoadPic_n
dim rsLayer
dim LastPostTimes
dim toptopic
dim TotalUseTable
dim FoundTable
dim ihaveupfile,upfileinfo,upfilelen
dim locktopic
dim ArticleRandom
dim speak

FoundTable=false

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if request("followup")="" then
	foundErr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
elseif not isInteger(request("followup")) then
	foundErr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	announceid=request("followup")
	ParentID=request("followup")
end if
if request("RootID")="" then
	foundErr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
elseif not isInteger(request("RootID")) then
	foundErr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	rootID=request("RootID")
end if

UserName=Checkstr(trim(request.Form("username")))
if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
else
	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
end if


IP=replace(Request.ServerVariables("REMOTE_ADDR"),"'","")
Expression=Checkstr(Request.Form("Expression"))
PostTopic=Checkstr(request.Form("subject"))
Body=Checkstr(request.Form("Content"))
signflag=Checkstr(trim(request.Form("signflag")))
mailflag=Checkstr(trim(request.Form("emailflag")))
msgflag=Checkstr(trim(request.Form("msgflag")))
ArticleRandom=Checkstr(trim(request.Form("ArticleRandom")))
TotalUseTable=Checkstr(request.Form("TotalUseTable"))
speak=Checkstr(trim(request.Form("speak")))

if Expression="" then
Expression=Forum_PostFace(0)
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
For i=0 to ubound(AllPostTable)
	if AllPostTable(i)=TotalUseTable then
		FoundTable=true
		Exit For
	end if
Next
if Not FoundTable then
	ErrMsg=ErrMsg+"<Br>"+"<li>您指定了非法的数据表名，请确认您是从有效的表单提交。"
	FoundErr=True
end if
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
	end if
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
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入姓名"
	foundErr=True
end if
if strLength(PostTopic)>50 then
	foundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>主题长度不能超过50"
end if
if request("method")="Topic" then
	if PostTopic="" then
		if body="" then
			ErrMsg=ErrMsg+"<Br>"+"<li>主题和内容必须填写其一。"
			foundErr=True
		end if
	end if
end if
if request("method")="fastreply" then
	if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>快速回复请填写发言内容。"
	foundErr=True
	end if
end if
if strLength(body)>Clng(Board_Setting(16)) then
	ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于" & CSTR(Board_Setting(16)) & "bytes"
	FoundErr=true
end if
if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>没有填写内容。"
	foundErr=true
end if
session("lastpost")=Now()
if founderr then
	stats="错误信息"
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats="回复帖子成功"
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
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发帖/回帖。"
	founderr=true
	exit sub
end if
if Cint(GroupSetting(5))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛回复其他人帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
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

sql="select locktopic,LastPost,title,postusername from topic where topicid="&cstr(rootid)
set rs=conn.execute(sql)
if not rs.eof and not rs.bof then
	toptopic=rs(2)
	if not master and not superboardmaster and rs("locktopic")=1 then
		Errmsg=ErrMsg+"<Br>"+"<li>本主题已经锁定，不能发表回复。"
		founderr=true
		exit sub
	elseif rs("locktopic")=2 then
		Errmsg=ErrMsg+"<Br>"+"<li>本主题已经删除，不能发表回复。"
		founderr=true
		exit sub
	end if
	dim rst
	sql="select top 1 f_id from [Friend] where F_username='"&rs(3)&"' and F_friend='"&membername&"' and F_type=1"
	set rst=conn.execute(sql)
	if not(rst.eof and rst.bof) then
		Errmsg=ErrMsg+"<Br>"+"<li>您已被此用户屏蔽，不能对此用户的帖子发表回复。"
		founderr=true
		exit sub
	end if
	rst.close	
	if not isnull(rs(1)) then
		LastPost=split(rs(1),"$")
		if ubound(LastPost)=7 then
		UpLoadPic_n=LastPost(4)
		else
		UpLoadPic_n=""
		end if
	end if
end if
set rs=nothing

usercookies=request.Cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies=3

if chkuserlogin(username,userpassword,usercookies,3)=false then
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

select case cint(forum_setting(73))
case 2
	if myUserPPD*cint(forum_setting(74))<=myTodayReply then
		Errmsg=ErrMsg+"<Br>"+"<li>您今天发表的回复数(<b>"&myTodayReply&"</b>)已经超过了限制(<b>"&(myUserPPD*cint(forum_setting(74)))&"</b>)，不能继续发表回复！"
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

set rsLayer=conn.execute("select b.layer,b.orders,b.EmailFlag,b.username,u.userEmail,b.topic,b.body from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.Announceid="&ParentID)
if not(rsLayer.eof and rsLayer.bof) then
	if isnull(rsLayer(0)) then
		iLayer=1
	else
		iLayer=rslayer(0)
	end if
	if isNUll(rslayer(1)) then
		iOrders=0
	else
		iOrders=rsLayer(1) 
	end if
	dim title
	if rsLayer("topic")="" or isnull(rsLayer("topic")) then
		if len(rsLayer("body"))>35 then
			Title=left(htmlencode(rsLayer("body")),35)+"..."
		else
			Title=htmlencode(rsLayer("body"))
		end if
	else
		if len(rsLayer("Topic"))>35 then
			Title=left(htmlencode(rsLayer("Topic")),35)+"..."
		else
			Title=htmlencode(rsLayer("Topic"))
		end if
	end if
	if rsLayer(3)=membername then
		if Cint(GroupSetting(4))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛回复自己主题的权限。"
		founderr=true
		exit sub
		end if
	end if
	if (rsLayer(2) mod 2)=1 and lcase(trim(username))<>lcase(trim(rsLayer(3))) then
		topic="您在"&Forum_info(0)&"发表的文章有人回复了"
		email=rsLayer(4)
		mailbody=rsLayer(3)&"，您好：<br>"
		mailbody=mailbody & "　　<Font color=blue>"&username&"</font>于<b>"&now&"</b>回复了您在"&Forum_info(0)&"发表的文章<font color=red>《"&Title&"》</font>，"
		mailbody=mailbody & "请到以下地址浏览该帖子内容。<br>"
		mailbody=mailbody & "<a href=http://"&request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),"savereannounce.asp","")&"dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&parentid&" target=_blank>查看帖子内容</a>"
		if Forum_Setting(2)=0 then
                               
		elseif Forum_Setting(2)=1 then
			call jmail(email,topic,mailbody)
		elseif Forum_Setting(2)=2 then
			call Cdonts(email,topic,mailbody)
		elseif Forum_Setting(2)=3 then
			call aspemail(email,topic,mailbody)
		end if
	end if
	if board_setting(46)=1 and rsLayer(2)>1 and lcase(trim(username))<>lcase(trim(rsLayer(3))) then
		topic="您在"&Forum_info(0)&"发表的文章有人回复了"
		email=rsLayer(3)
		mailbody=rsLayer(3)&"，您好："&chr(10)
		mailbody=mailbody & "　　[color=blue]"&username&"[/color]于[b]"&now&"[/b]回复了您在"&Forum_info(0)&"发表的文章[color=red]《"&Title&"》[/color]，"
		mailbody=mailbody & "请到以下地址浏览该帖子内容。"&chr(10)
		mailbody=mailbody & "[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&parentid&"][color=blue][b][u]查看帖子内容[/u][/b][/color][/URL][/align]"
		if Board_Setting(46)=1 then
			dim rs2,sql2
			set rs2=server.createobject("adodb.recordset")
			sql2="select * from [message]"      
			rs2.open sql2,conn,1,3      
			rs2.addnew      
			rs2("sender")=forum_info(0)
			rs2("incept")=email
			rs2("title")=topic
			rs2("content")=mailbody
			rs2("flag")=0 
			rs2("issend")=1
			rs2.update
			rs2.close
		end if
	end if
else
	iLayer=1
	iOrders=0
end if
rsLayer.close
set rsLayer=nothing
if rootid<>0 then 
	iLayer=ilayer+1
	conn.execute "update "&TotalUseTable&" set orders=orders+1 where rootid="&cstr(RootID)&" and orders>"&cstr(iOrders)
	iOrders=iOrders+1
end if
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
DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"上午",""),"下午","")
'插入回复表
Sql="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,isaudit,speak) values ("&boardid&","&ParentID&",'"&username&"','"&PostTopic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&RootID&","&ilayer&","&iorders&",'"&ip&"','"&Expression&"',"&locktopic&","&signflag&","&(mailflag+msgflag*2)&",0,"&userid&","&ihaveupfile&","&isaudit&",'"&speak&"')"
conn.execute(sql)
set rs=conn.execute("select max(announceid) from "&TotalUseTable&"")
announceid=rs(0)

if ihaveupfile=1 then conn.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&replace(toptopic,"'","")&"' where F_ID in ("&upfileinfo&")")

rs.close
set rs=nothing
conn.execute("update [user] set userwealth=userwealth-"&(msgusemoney+mailusemoney)&" where username='"&membername&"'")
if PostTopic="" then
	PostTopic=replace(cutStr(reUBBCode(body,false),14),chr(10),"")
else
	PostTopic=replace(cutStr(reUBBCode(PostTopic,false),14),chr(10),"")
end if
LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(PostTopic,"$","") & "$" & UpLoadPic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost=replace(LastPost,"'","")
'更新主题表记录
if isaudit=0 then
	sql="update topic set child=child+1,LastPostTime=Now(),LastPost='"&LastPost&"' where TopicID="&rootID
	conn.execute(sql)
else
	conn.execute("update "&TotalUseTable&" set isaudit=1 where rootid="&rootid)
end if

toptopic=replace(replace(cutStr(reUBBCode(toptopic,false),14),chr(10),""),"'","")
LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(topTopic,"$","") & "$" & UpLoadPic_n & "$" & UserID & "$" & rootid & "$" & BoardID
'回复首页显示主题
'LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(PostTopic,"$","") & "$" & UpLoadPic_n & "$" & UserID & "$" & rootid & "$" & BoardID  '主题显示回复
LastPost_1=replace(LastPost_1,"'","")
if isaudit=0 then
Dim UpdateBoardID
UpdateBoardID=BoardParentStr & "," & BoardID
if datediff("d",LastPostTime,Now())=0 then
	sql="update board set lastbbsnum=lastbbsnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
else
	sql="update board set lastbbsnum=lastbbsnum+1,todaynum=1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
end if
conn.execute(sql)

Dim updateinfo
set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
LastPostTimes=split(rs(0),"$")
LastPostTime=LastPostTimes(2)
if not isdate(LastPostTime) then LastPostTime=Now()
if datediff("d",LastPostTime,Now())=0 then
if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
	sql="update config set bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost_1&"' "&updateinfo&" where active=1"
else
	sql="update config set bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost_1&"' where active=1"
end if
conn.execute(sql)
end if
		
call haveRe()
response.cookies("LastView")("Topic_"&RootID&"_"&BoardID)=now()

call nav()
call head_var(1,BoardDepth,0,0)
dim PostRetrunName
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
	response.write "<meta http-equiv=refresh content=""5;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&Announceid&""">"
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