<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="urlconn.asp"-->

<%
dim act
stats="在线收藏夹"
call nav()
call head_var(2,0,"","")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本社区在线收藏夹的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call dvbbs_error()
else
	response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=20><b>在线收藏夹</b></td><tr><td valign=middle class=tablebody1>"
	act=request("act")
	select case act
	case "add"
		call main()
	case "del"
		call del()
	case "edit"
		call edit()
	case "added"
		call added()
	case "edited"
		call edited()
	case "out"
		call out()
	case else
		response.write "<script>alert('请从正常渠道进入！');history.go(-1);</script>"
		response.end
	end select
end if

response.write "</table>"
call footer()
%>

<%
sub main()
	dim group,userid
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
	else
		group=rs("groups")
		userid=rs("id")
		response.write "<form name=thisform method=post action=?act=added><table cellpadding=3 cellspacing=1 align=center class=tableborder1 align=center>"
		response.write "<tr><td class=tablebody2 align=center><b>添加组</b>"
		response.write "<tr><td class=tablebody1>已经存在的分组："
		for i=0 to rs("groupnum")-1
			response.write "&nbsp;&nbsp;<b>"&i+1&"、</b>"&trim(split(group,"|")(i))
		next
		response.write "<tr><td class=tablebody2 align=center>"
		response.write "请输入要添加的组名：<input type=text name=groupname size=4> <input type=submit value=添加>"
		response.write "</table></form>"
	end if
end sub

sub edit()
	dim group,userid,tmpname
	tmpname=trim(request("name"))
	if tmpname="" then
		response.write "<script>alert('请从正常渠道进入！');history.go(-1);</script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
	else
		group=rs("groups")
		userid=rs("id")
		response.write "<form name=thisform method=post action=?act=edited><table cellpadding=3 cellspacing=1 align=center class=tableborder1 align=center>"
		response.write "<tr><td class=tablebody2 align=center><b>组改名</b>"
		response.write "<tr><td class=tablebody1>已经存在的分组："
		for i=0 to rs("groupnum")-1
			response.write "&nbsp;&nbsp;<b>"&i+1&"、</b>"&trim(split(group,"|")(i))
		next
		response.write "<tr><td class=tablebody2 align=center>原组名：<font color=red>"&tmpname&"</font><input type=hidden name=oldname value="&tmpname&">"
		response.write "<tr><td class=tablebody1 align=center>"
		response.write "请输入修改后的组名：<input type=text name=groupname size=4> <input type=submit value=修改>"
		response.write "</table></form>"
	end if
end sub

sub added()
	dim groupname,group,tmpstr
	groupname=trim(request("groupname"))
		if groupname="" then
			response.write "<script>alert('请输入您要添加的组名！');history.go(-1);</script>"
			response.end
		elseif instr(groupname,"|") then
			response.write "<script>alert('您的输入存在非法字符,如|等！');history.go(-1);</script>"
			response.end
		else
			set rs=server.createobject("adodb.recordset")
			sql="select * from [fav] where username='"&membername&"'"
			rs.open sql,connurl,1,3
			if rs.eof and rs.bof then
				response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
			else
				group=rs("groups")&"|"
				tmpstr="|" & groupname & "|"
				if instr(group,tmpstr) then
						response.write "<script>alert('您输入的组名已经存在！');history.go(-1);</script>"
						response.end
				end if			
				rs("groups")=group & groupname
				rs("groupnum")=rs("groupnum")+1
				rs.update
				response.redirect "url.asp"
			end if
		end if
end sub


sub edited()
	dim groupname,group,oldname,tmpstr
	groupname=trim(request("groupname"))
	oldname=trim(request("oldname"))
		if groupname="" then
			response.write "<script>alert('请输入修改后的组名！');history.go(-1);</script>"
			response.end
		elseif instr(groupname,"|") then
			response.write "<script>alert('您的输入存在非法字符,如|等！');history.go(-1);</script>"
			response.end
		elseif oldname="" then
			response.write "<script>alert('请从正常渠道进入!!！');history.go(-1);</script>"
			response.end
		else
			set rs=server.createobject("adodb.recordset")
			sql="select * from [fav] where username='"&membername&"'"
			rs.open sql,connurl,1,3
			if rs.eof and rs.bof then
				response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
			else
				group=rs("groups")
				tmpstr="|" & groupname & "|"
				if instr(group&"|",tmpstr) then
						response.write "<script>alert('您输入的组名已经存在！');history.go(-1);</script>"
						response.end
				end if			
				rs("groups")=replace(group,oldname,groupname)
				rs.update
				rs.close
				sql="select * from [url] where username='"& membername &"' and grouptype='"& oldname &"'"
				rs.open sql,connurl,1,3
				if not(rs.eof and rs.bof) then
					rs.movefirst
					do while not rs.eof
						rs("grouptype")=groupname
						rs.update
						rs.movenext
					loop
				end if
				response.redirect "url.asp"
			end if
		end if
end sub

sub del()
	dim group,userid,tmpname,tmpstr,groupid
	tmpname=trim(request("name"))
	if tmpname="" then
		response.write "<script>alert('请从正常渠道进入！');history.go(-1);</script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,3
	if rs.eof and rs.bof then
		response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
	else
		group=rs("groups")
		userid=rs("id")
		for i=0 to rs("groupnum")-1
			if tmpname<>trim(split(group,"|")(i)) then tmpstr=tmpstr & trim(split(group,"|")(i)) & "|"
		next
		tmpstr=mid(tmpstr,1,len(tmpstr)-1)
		rs("groups")=tmpstr
		rs("groupnum")=rs("groupnum")-1
		rs.update
		rs.close
		sql="delete from url where username='"&membername&"' and grouptype='"&tmpname&"'"
		rs.open sql,connurl,1,3
		response.redirect "url.asp"
	end if
end sub

sub out()
	set rs=server.createobject("adodb.recordset")
	sql="delete from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,3
	sql="delete from url where username='"&membername&"'"
	response.redirect "url.asp"
end sub

%>