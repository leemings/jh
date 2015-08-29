<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<% 
dim announceid
dim UserName
dim userPassword
dim useremail
dim Topic
dim body
dim ip
dim Expression
dim rootid
dim signflag
dim mailflag
dim msgflag
dim char_changed
dim addbody
dim totalusetable
Dim FoundTable
dim ChenTopicColor
dim speak

FoundTable=false
addbody=checkStr(request("Content"))
dim re
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="\[align=right\]\[color=#000066\](.*)\[\/color\]\[\/align\]"
addbody=re.Replace(addbody,"")
set re=Nothing

if membername<>request("username") then 
	if Cint(forum_setting(49))=1 then
		char_changed = chr(13) & chr(10) & "[align=right][color=#000066][此帖子已经被"&membername&"于"&Now()&"编辑过][/color][/align]" & chr(13)
	end if
else
	if not (master and Cint(forum_setting(49))=0) and Cint(forum_setting(48))=1 then
		char_changed = chr(13) & chr(10) & "[align=right][color=#000066][此帖子已经被作者于"&Now()&"编辑过][/color][/align]" & chr(13)
	end if
end if

UserName=trim(checkStr(request("username")))
UserPassWord=trim(checkStr(request("passwd")))
IP=replace(Request.ServerVariables("REMOTE_ADDR"),"'","")
Expression=checkStr(Request.Form("Expression"))
BoardID=checkStr(Request("boardID"))
rootID=Cstr(checkStr(Request("ID")))
AnnounceID=checkStr(request("replyID"))
Topic=checkStr(request("subject"))
signflag=trim(checkStr(request("signflag")))
mailflag=trim(checkStr(request("emailflag")))
msgflag=trim(checkStr(request("msgflag")))
ChenTopicColor=Checkstr(trim(request.form("ChenTopicColor")))
speak=Checkstr(trim(request.form("speak")))
foundErr=false
TotalUseTable=Checkstr(request.Form("TotalUseTable"))

For i=0 to ubound(AllPostTable)
	if AllPostTable(i)=TotalUseTable and not isnull(TotalUseTable) then
		FoundTable=true
		Exit For
	end if
Next
if Not FoundTable then
	ErrMsg=ErrMsg+"<Br>"+"<li>您指定了非法的数据表名，请确认您是从有效的表单提交。"
	FoundErr=True
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
	call activeonline()
	call footer()
	response.end
end if
if not founduser then
	ErrMsg=ErrMsg+"<Br>"+"<li>请登录后进行修改。"
	foundErr=True
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
if instr(Expression,"face")=0 then
	Expression="face1.gif"
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>您提交的数据不合法，请不要从外部提交发言。"
   	FoundErr=True
end if
if AnnounceID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(AnnounceID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
end if
if BoardID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(BoardID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
end if
if rootID="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(rootID) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入姓名(长度不能大于20)"
 	foundErr=True
elseif Trim(UserPassWord)="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>请输入密码(长度不能大于16)"
   	foundErr=True
end if
''''''''''''''''''''''''''''''''''
'判断主题是否为空.
set rs=conn.execute("select announceid from "&TotalUseTable&" where  ParentID=0 and rootid="&rootid&" ")
if clng(announceid)=clng(rs(0)) then
	if Topic="" then
   		foundErr=True
		ErrMsg=ErrMsg+"<Br>"+"<li>主题不应为空"
	elseif strLength(topic)>50 then
   		foundErr=True
      	        ErrMsg=ErrMsg+"<Br>"+"<li>主题长度不能超过50"
	end if
end if
rs.close:set rs=nothing
'''''''''''''''''''''''''''''''''''
if request("content")="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>内容不能为空"
 	foundErr=True
end if
if strLength(body)>Clng(Board_Setting(16)) then
	ErrMsg=ErrMsg+"<Br>"+"<li>发言内容不得大于" & CSTR(Board_Setting(16)) & "bytes"
 	foundErr=true
end if

if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发帖。"
	founderr=true
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
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

stats="编辑帖子成功"

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
call activeonline()
call footer()

sub main()
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
dim LastBoard
dim LastTopic
dim LastPost
dim caneditpost
caneditpost=false
sql="select b.username,b.dateandtime,u.usergroupID from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.rootid="&rootid&" and b.AnnounceID="&AnnounceID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>没有找到相应的帖子。"
	Founderr=true
	exit sub
else
	if Clng(forum_setting(50))>0 then
		if Datediff("s",rs("dateandtime"),Now())>Clng(forum_setting(50))*60 then
			Body=addbody+char_changed
		else
			Body=addbody
		end if
	else
		Body=addbody+char_changed
	end if
	if rs("username")=membername then
		if Cint(GroupSetting(10))=0 then
			Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛编辑自己帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
			founderr=true
			CanEditPost=False
		else
			CanEditPost=True
		end if
	else
		if (master or superboardmaster or boardmaster) and Cint(GroupSetting(23))=1 then
			CanEditPost=True
		else
			CanEditPost=False
		end if
		if UserGroupID>3 and Cint(GroupSetting(23))=1 then
			CanEditPost=true
		end if
		if Cint(GroupSetting(23))=1 and FoundUserPer then
			CanEditPost=True
		elseif Cint(GroupSetting(23))=0 and FoundUserPer then
			CanEditPost=False
		end if
		if UserGroupID<3 and UserGroupID=rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>同等级用户不能修改。"
			Founderr=true
			exit sub
		elseif UserGroupID=3 and UserGroupID=rs("UserGroupID") then
			if boardmaster then
				dim rss,userboardmaster
				sql="select BoardMaster from board where boardid="&boardid
				set rss=server.createobject("adodb.recordset")
				rss.open sql,conn,1,1
				userboardmaster=rss(0)
				rss.close
				set rss=nothing
				if instr(1,"|"&lcase(trim(userboardmaster))&"|","|"&lcase(trim(rs("username")))&"|")>0 then
					Errmsg=Errmsg+"<br>"+"<li>同等级用户不能修改。"
					Founderr=true
					exit sub
				end if
			end if
		elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>不能修改等级比您高的用户的帖子。"
			Founderr=true
			exit sub
		end if
		if not CanEditPost then
			Errmsg=Errmsg+"<br>"+"<li>您没有足够的权限编辑本帖子，请和管理员联系。"
			Founderr=true
			exit sub
		end if
	end if
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT * FROM "&TotalUseTable&" where username='"&trim(username)&"' and AnnounceID="&Announceid
rs.Open sql,conn,1,3
if rs.eof and rs.bof then
	foundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>您不是本帖子的作者，无权修改！"
	rs.close:set rs=nothing
	exit sub
elseif not master and not superboardmaster and rs("locktopic")=1 then
	Errmsg=ErrMsg+"<Br>"+"<li>本主题已经锁定，不能编辑。"
	foundErr=true
	rs.close:set rs=nothing
	exit sub
else
	dim mailusemoney,msgusemoney
	mailusemoney=0:msgusemoney=0
	if rs("emailflag")<>mailflag+msgflag*2 then
		if rs("emailflag") \ 2 <> msgflag and msgflag=1 then
			if isnull(myvip) or myvip<>1 then
				msgusemoney=clng(forum_user(28))
			else
				msgusemoney=clng(forum_user(29))
			end if
		end if
		if rs("emailflag") mod 2 <> mailflag and mailflag=1 then
			if isnull(myvip) or myvip<>1 then
				mailusemoney=clng(forum_user(30))
			else
				mailusemoney=clng(forum_user(31))
			end if
		end if
		if mymoney<msgusemoney+mailusemoney then
			Errmsg=ErrMsg+"<Br>"+"<li>您的现金不够支付短信通知或邮件通知的费用！"
			founderr=true
			rs.close:set rs=nothing
			exit sub
		end if
	end if
	rs("Topic") =replace(Topic,"''","'")
	rs("Body") =replace(Body,"''","'")
	rs("length")=strlength(body)
	rs("ip")=ip
	rs("Expression")=Expression
	rs("signflag")=signflag
	rs("emailflag")=(mailflag+msgflag*2)
	rs("speak")=speak
	rs.Update

	if rs("parentid")=0 then
		conn.execute("update topic set title='"&topic&"',LastPostTime=Now(),ChenTopicColor='"&ChenTopicColor&"' where topicid="&rootid)
	end if
	conn.execute("update [user] set userwealth=userwealth-"&(msgusemoney+mailusemoney)&" where username='"&membername&"'")

	rs.close:set rs=nothing
	'取出当前版面最后回复id,如果本帖为最后回复则更新相应数据
	sql="select LastPost from board where boardid="&boardid
	set rs=conn.execute(sql)
	if not (rs.eof and rs.bof) then
		if not isnull(rs(0)) and rs(0)<>"" then
			LastBoard=split(rs(0),"$")
			if ubound(LastBoard)=7 then
				if Clng(LastBoard(6))=Clng(AnnounceID) then
				LastPost=LastBoard(0) & "$" & LastBoard(1) & "$" & Now() & "$" & replace(cutStr(reUBBCode(Topic,false),20),"$","") & "$" & LastBoard(4) & "$" & LastBoard(5) & "$" & LastBoard(6) & "$" & boardid
				conn.execute("update board set LastPost='"&replace(LastPost,"'","")&"' where boardid="&boardid)
				end if
			end if
		end if
	end if
	
	'取得当前主题最后回复id,如果本帖为最后回复则更新相应数据
	sql="select LastPost from topic where topicid="&rootid
	set rs=conn.execute(sql)
	if not (rs.eof and rs.bof) then
		if not isnull(rs(0)) and rs(0)<>"" then
			LastTopic=split(rs(0),"$")
			if ubound(LastTopic)=7 then
				if Clng(LastTopic(1))=Clng(Announceid) then
				LastPost=LastTopic(0) & "$" & LastTopic(1) & "$" & Now() & "$" & replace(cutStr(reUBBCode(body,false),20),"$","") & "$" & LastTopic(4) & "$" & LastTopic(5) & "$" & LastTopic(6) & "$" & boardid
				conn.execute("update topic set LastPost='"&replace(LastPost,"'","")&"' where topicid="&rootid)
				end if
			end if
		end if
	end if
	response.cookies("LastView")("Topic_"&RootID&"_"&BoardID)=now()

	call nav()
	call head_var(1,BoardDepth,0,0)
	dim PostRetrunName
	select case Board_Setting(17)
	case 1
		response.write "<meta http-equiv=refresh content=""1;URL=index.asp"">"
		PostRetrunName="首页"
	case 2
		response.write "<meta http-equiv=refresh content=""1;URL=list.asp?boardid="&boardid&""">"
		PostRetrunName="您所发布的论坛"
	case 3
		response.write "<meta http-equiv=refresh content=""1;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&"&star="&request("star")&"#"&rootid&""">"
		PostRetrunName="您所修改的帖子"
	end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">状态：<%=stats%></th>
</tr><tr><td width="100%" class=tablebody1>
本页面将在1秒后自动返回<%=PostRetrunName%>，<b>您可以选择以下操作：</b><br><ul>
<li><a href="index.asp">返回首页</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a></li>
</ul></td></tr></table>
<%
end if
end sub
%>