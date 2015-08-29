<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
	response.buffer=true
	dim rootid
	dim AnnounceID
	stats="购买帖子"
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
	else
		rootid=request("id")
	end if
	if request("replyID")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
	elseif not isInteger(request("replyID")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
	else
		AnnounceID=request("replyID")
	end if
	dim FoundTable
	FoundTable=false
	if request("PostTable")<>"" then
	For i=0 to ubound(AllPostTable)
		if request("PostTable")=AllPostTable(i) then
			FoundTable=true
			exit for
		end if
	next
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	end if
	if not FoundTable then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	end if
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
	end if

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,boarddepth,"","")
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
	dim re,test
	dim po,ii,jj
	dim reContent
	dim strContent
	dim PostBuyUser
	dim Title,Topic,Body
	dim rs2,sql2
	po=0
	ii=0
	jj=0
	dim usermoney
	set rs=conn.execute("select userWealth from [user] where userid="&userid)
	usermoney=rs(0)
	set rs=server.createobject("adodb.recordset")
	sql="select body,PostBuyUser,username,PostUserID,rootId,Topic from "&request("PostTable")&" where Announceid="&Announceid
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>错误的帖子参数。"
		exit sub
	else
		strContent=HTMLcode(rs(0))
		PostBuyUser=rs(1)
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		re.Pattern="\[USEMONEY=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USEMONEY\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USEMONEY=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "USEMONEY=$1" & chr(2))
				re.Pattern="\[\/USEMONEY\]"
				strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
				re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(\x01\/USEMONEY\x02)"
				strContent=re.Replace(strContent,"")
				re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
				po=re.Replace(strContent,"$3")
				if IsNumeric(po) then
					ii=int(po) 
				else
					ii=0
				end if
			end if
		else
			re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[\/USEMONEY\]"
				Test=re.Test(strContent)
				if Test then
					re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
					strContent=re.replace(strContent, chr(1) & "USEMONEY=$1,$2" & chr(2))
					re.Pattern="\[\/USEMONEY\]"
					strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
					re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(\x01\/USEMONEY\x02)"
					strContent=re.Replace(strContent,"")
					re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*),([0|1])\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					po=re.Replace(strContent,"$4")
					if IsNumeric(po) then
						jj=int(po) 
					else
						jj=0
					end if
				end if
			end if
		end if
		if membername=rs(2) then
			response.write "<script>alert('呵呵，您要花钱购买自己发布的帖子嘛？');</script>"
		elseif usermoney>ii then
			if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
				if instr("|"&PostBuyUser&"|","|"&membername&"|")>0 then
					response.write "<script>alert('呵呵，您已经购买过了呀？');</script>"
				else
					conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
					conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
					rs(1)=rs(1) & "|" & membername
					rs.update
					if rs("topic")="" or isnull(rs("topic")) then
						if len(rs("body"))>35 then
							Title=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),35)+"..."
						else
							Title=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
						end if
					else
						if len(rs("Topic"))>35 then
							Title=left(htmlencode(rs("Topic")),35)+"..."
						else
							Title=htmlencode(rs("Topic"))
						end if
					end if
					if jj=1 then
						Topic="您在"&Forum_info(0)&"出售的文章有人购买了"
						body=rs("username")&"，您好："&chr(10)
						body=body & "　　[color=blue]"&membername&"[/color]于[b]"&now&"[/b]购买了您在"&Forum_info(0)&"发表的文章[color=red]《"&Title&"》[/color]，"
						body=body & "请到以下地址浏览该帖子内容。"&chr(10)
						body=body & "[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rs("rootid")&"&replyid="&announceid&"&skin=1][color=blue][b][u]查看帖子内容[/u][/b][/color][/URL][/align]"
						set rs2=server.createobject("adodb.recordset")
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")=forum_info(0)      
						rs2("incept")=rs("username")
						rs2("title")=topic
						rs2("content")=body
						rs2("flag")=0 
						rs2("issend")=1
						rs2.update
						rs2.close
						set rs2=nothing
					end if
				end if
				response.write "<script>alert('购买成功！');</script>"
			else
				conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
				conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
				rs(1)=membername
				rs.update
				if rs("topic")="" or isnull(rs("topic")) then
					if len(rs("body"))>35 then
						Title=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),35)+"..."
					else
						Title=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
					end if
				else
					if len(rs("Topic"))>35 then
						Title=left(htmlencode(rs("Topic")),35)+"..."
					else
						Title=htmlencode(rs("Topic"))
					end if
				end if
				if jj=1 then
					Topic="您在"&Forum_info(0)&"出售的文章有人购买了"
					body=rs("username")&"，您好："&chr(10)
					body=body & "　　[color=blue]"&membername&"[/color]于[b]"&now&"[/b]购买了您在"&Forum_info(0)&"发表的文章[color=red]《"&Title&"》[/color]，"
					body=body & "请到以下地址浏览该帖子内容。"&chr(10)
					body=body & "[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rs("rootid")&"&replyid="&announceid&"&skin=1][color=blue][b][u]查看帖子内容[/u][/b][/color][/URL][/align]"
					set rs2=server.createobject("adodb.recordset")
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")=forum_info(0)      
					rs2("incept")=rs("username")
					rs2("title")=topic
					rs2("content")=body
					rs2("flag")=0 
					rs2("issend")=1
					rs2.update
					rs2.close
					set rs2=nothing
				end if
				response.write "<script>alert('购买成功！');</script>"
			end if
		else
			response.write "<script>alert('您都没有钱呀？');</script>"
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid&"&replyID="&AnnounceID&"&skin=1"
end sub
%>