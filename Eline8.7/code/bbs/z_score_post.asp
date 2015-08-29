<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
response.buffer=true
dim rootid,isScore,AnnisScore,CanScore,replyid
dim TotalUseTable
stats="帖子评分"
if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if cint(Board_Setting(49))=0 then
	Errmsg=Errmsg+"<br>"+"<li>这个版面不允许版主评分。"
	founderr=true
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	rootid=clng(request("id"))
end if
if request("replyid")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(request("replyid")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	replyid=clng(request("replyid"))
end if
if request("rate")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关参数。"
elseif not isNumeric(request("rate")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
else
	isScore=cint(request("rate"))
	if isScore>5 or isScore<-5 or isScore<>int(isScore) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	end if
end if
'判断用户是否有评分的权限
if master or superboardmaster or boardmaster then
	canScore=true
else
	canScore=false
end if
if not canScore then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛评分的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
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
call activeonline()
call footer()

sub main()
	set rs=server.createobject("adodb.recordset")
	set rs=conn.execute("select PostTable from topic where topicid="&rootid)
	TotalUseTable=rs(0)
	rs.close
	sql="select score,scoreuser,scoretime,announceid,boardid,username,topic,body from "&TotalUseTable&" where announceid="&replyid
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>错误的帖子参数。"
		exit sub
	else
		select case isCanScore(rs("scoreuser"),rs("username"))
		case 3
			Errmsg=Errmsg+"<br>"+"<li>您不能给自己评分。"
			founderr=true
			exit sub
		case 4
			Errmsg=Errmsg+"<br>"+"<li>您不能修改相同用户等级的评分。"
			founderr=true
			exit sub
		case 5
			Errmsg=Errmsg+"<br>"+"<li>您不能修改更高用户等级的评分。"
			founderr=true
			exit sub
		end select
		if not isnull(rs("score")) then
			if isScore=rs("score") then
				Errmsg=Errmsg+"<br>"+"<li>请不要重复评分。"
				founderr=true
				exit sub
			end if
		end if
	end if
	dim topic
	if rs("topic")="" or isnull(rs("topic")) then
		if len(rs("body"))>20 then
			topic=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),20)+"..."
		else
			topic=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
		end if
	else
		if len(rs("Topic"))>20 then
			topic=left(htmlencode(rs("Topic")),20)+"..."
		else
			topic=htmlencode(rs("Topic"))
		end if
	end if
	if trim(topic)="" then topic="..."
	if isnull(rs("score")) then
		AnnisScore=isScore
	else
		AnnisScore=isScore-rs("score")
	end if
	rs("score")=isScore
	rs("scoreuser")=membername
	rs("scoretime")=now()
	rs.update
	'修改用户的积分
	conn.execute("update [user] set userScore=userScore+("&AnnisScore&") where username='"&rs("username")&"'")
	'记录论坛事件
	if AnnisScore>0 then
		conn.execute("insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&rs("announceid")&","&rs("boardid")&",'"&rs("username")&"','"&membername&"','给《"&topic&"》加了"&AnnisScore&"分','"&Request.ServerVariables("REMOTE_ADDR")&"')")
	elseif AnnisScore<0 then
		conn.execute("insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&rs("announceid")&","&rs("boardid")&",'"&rs("username")&"','"&membername&"','给《"&topic&"》减了"&abs(AnnisScore)&"分','"&Request.ServerVariables("REMOTE_ADDR")&"')")
	end if
	rs.close
	set rs=nothing
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&request("id")&"&star="&request("star")&"#"&request("replyid")
end sub
%>