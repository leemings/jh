<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
	response.buffer=true
	dim rootid,isagree,Annisagree,CanAgree
	dim getmoney,TotalUseTable
	'设置投票所用金钱
	getmoney=Cint(GroupSetting(47))
	stats="帖子投票"
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	'判断用户是否有鲜花/鸡蛋权限
	if Cint(GroupSetting(6))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表看法的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
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
	
	if request.cookies("aspsky")("userid")=request("userid") then	
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>你不能给自己投票。"
	end if
        
	if session("postagree")<>"" and not master then
		if clng(session("postagree"))=clng(rootid) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>20分钟内只能投一次。"
		end if
	end if
	if request("isagree")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关参数。"
	elseif not isInteger(request("isagree")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	else
		isagree=request("isagree")
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
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
	dim topic
	if mymoney<getmoney then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您没有足够的金钱发表看法。"
		exit sub
	else
		conn.execute("update [user] set userWealth=userWealth-"&getmoney&" where userid="&userid)
		set rs=server.createobject("adodb.recordset")
		set rs=conn.execute("select PostTable,Title from topic where topicid="&rootid)
		TotalUseTable=rs(0)
		topic=rs(1)
		rs.close
		sql="select top 1 isagree,announceid,boardid,username from "&TotalUseTable&" where rootid="&rootid&" order by Announceid"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>错误的帖子参数。"
			exit sub
		else
			dim tempisagree
			tempisagree=isagree
			if not isnull(rs(0)) and rs(0)<>"" then
				Annisagree=split(rs(0),"|")
				if Cint(isagree)=1 then
					isagree=Annisagree(0)+1
					rs("isagree")=isagree & "|" & Annisagree(1)
				else
					isagree=Annisagree(1)+1
					rs("isagree")=Annisagree(0) & "|" & isagree
				end if
				rs.update
			else
				if Cint(isagree)=1 then
					rs("isagree")="1|0"
				else
					rs("isagree")="0|1"
				end if
				rs.update
			end if
			'修改用户的威望值	
			if Cint(tempisagree)=1 then conn.execute("update [user] set userPower=userPower+1 where userid="&request("userid"))
			if Cint(tempisagree)=2 then conn.execute("update [user] set userPower=userPower-1 where userid="&request("userid"))
		end if
		
		rs.close
		set rs=nothing
	end if

	session("postagree")=rootid
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid
end sub
%>