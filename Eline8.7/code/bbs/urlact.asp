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
	case "move"
		call move()
	case else
		response.write "<script>alert('请选择您要进行的操作！');history.go(-1);</script>"
		response.end
	end select
end if

response.write "</table>"
call footer()

sub main()
%>
<script language=vbscript>
<!--
function chform(theform)
	if trim(theform.webname.value)="" then
		alert("请输入收藏地址的名称！")
	elseif trim(theform.weburl.value)="" then
		alert("请输入要收藏的地址")
	else
		theform.submit
	end if
end function
//-->
</script>
<%
	dim group,userid
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
	else
		group=rs("groups")
		userid=rs("id")
		response.write "<form name=thisform method=post action=?act=added onSubmit='chform(this);return false;'><table cellpadding=3 cellspacing=1 align=center valign=top class=tableborder1>"
		response.write "<tr><td class=tablebody2 align=center colspan=3><b>添加地址</b>"
		response.write "<tr><td class=tablebody1 align=left colspan=3>"
		response.write "&nbsp;&nbsp;请填写下表以完成地址的添加,<font color=red>*</font>为必需内容。"
		response.write "<tr><td width=20% class=tablebody2><td class=tablebody1 align=left>收藏名字：<input type=text name=webname>&nbsp;<font color=red>*</font>"
		response.write "<br>收藏地址：<input type=text name=weburl size=30 value=http://>&nbsp;<font color=red>*</font>"
		response.write "<br>所属&nbsp;&nbsp;组：<select name=group>"
		for i=0 to rs("groupnum")-1
			response.write "<option value="&trim(split(group,"|")(i))&">"&trim(split(group,"|")(i))
		next
		response.write "</select>&nbsp;<font color=red>*</font>请选择"
		response.write "<br>简单说明：<input type=text name=webtext>"
		response.write "<br><input type=reset value=重置>&nbsp;&nbsp;&nbsp;<input type=submit value=添加>"
		response.write "<td width=20% class=tablebody2></table></form>"
	end if
end sub

sub added()
	dim webname,weburl,group,webtext
	dim userid
	webname=trim(request("webname"))
	weburl=trim(request("weburl"))
	group=trim(request("group"))
	webtext=trim(request("webtext"))
		if webname="" or weburl="" or group="" then
			response.write "<script>alert('请从正常渠道进入！');history.go(-1);</script>"
			response.end
		end if
		set rs=server.createobject("adodb.recordset")
		sql="select * from [fav] where username='"&membername&"'"
		rs.open sql,connurl,1,1
		if rs.eof and rs.bof then
			response.write "<div align=center>您还没有在本社区建立个人收藏夹，马上<a href='urlnew.asp'><font color=red>申请</font></a>在线收藏夹。</div>"
		else
			userid=rs("id")
			rs.close
			sql="select * from [url] where (id is null)"
			rs.open sql,connurl,1,3
			rs.addnew
			rs("webname")=webname
			rs("weburl")=weburl
			rs("grouptype")=group
			rs("webtext")=webtext
			rs("username")=membername
			rs.update
			response.redirect "url.asp"
		end if
end sub

sub edit()
%>
<script language=vbscript>
<!--
function xxx(form1)
	if trim(form1.webname.value)="" then
		alert("请输入收藏地址的名称！")
	elseif trim(form1.weburl.value)="" then
		alert("请输入要收藏的地址")
	else
		form1.submit
	end if
end function
//-->
</script>
<%
	dim urlid
	urlid=request("urlid")
	if urlid="" then
		response.write "<script>alert('请从正当渠道进入！');history.go(-1);</script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from url where id="&urlid
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		response.write "<script>alert('请从正当渠道进入！！');history.go(-1);</script>"
		response.end
	else
		response.write "<form name=thisform1 method=post action=?act=edited onSubmit='xxx(this);return false;'><table cellpadding=3 cellspacing=1 align=center valign=top class=tableborder1>"
		response.write "<tr><td class=tablebody2 align=center colspan=3><b>编辑地址</b>"
		response.write "<tr><td class=tablebody1 align=left colspan=3>"
		response.write "&nbsp;&nbsp;请填写下表以完成地址的编辑,<font color=red>*</font>为必需内容。"
		response.write "<tr><td width=20% class=tablebody2><td class=tablebody1 align=left>收藏名字：<input type=text name=webname value="&rs("webname")&">&nbsp;<font color=red>*</font>"
		response.write "<input type=hidden name=urlid value="&urlid&">"
		response.write "<br>收藏地址：<input type=text name=weburl value="&rs("weburl")&" size=30>&nbsp;<font color=red>*</font>"
		response.write "<br>简单说明：<input type=text name=webtext value="&rs("webtext")&">"
		response.write "<br><input type=reset value=重置>&nbsp;&nbsp;&nbsp;<input type=submit value=修改>"
		response.write "<td width=20% class=tablebody2></table></form>"
	end if
	
end sub

sub edited()
	dim urlid,webname,weburl,webtext
	urlid=trim(request("urlid"))
	webname=trim(request("webname"))
	weburl=trim(request("weburl"))
	webtext=trim(request("webtext"))
	if webname="" or urlid="" or weburl="" then
		response.write "<script>alert('请从正常渠道进入！');history.go(-1);</script>"
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from [url] where id="&urlid
	rs.open sql,connurl,1,3
	if rs.eof and rs.bof then
		response.write "<script>alert('请从正当渠道进入！！');history.go(-1);</script>"
		response.end
	else
		rs("webname")=webname
		rs("weburl")=weburl
		rs("webtext")=webtext
		rs.update
		response.redirect "url.asp"
	end if
end sub

sub del()
	dim urlid
	urlid=request("acter")
	if request("chkall")=1 then
		set rs=server.createobject("adodb.recordset")
		sql="delete from [url] where username='"&membername&"'"
		rs.open sql,connurl,1,3
	else
		if urlid="" then
			response.write "<script>alert('请选择您要操作的收藏地址！');history.go(-1);</script>"
			response.end
		else
			set rs=server.createobject("adodb.recordset")
			sql="delete from [url] where username='"&membername&"' and id in ("&urlid&")"
			rs.open sql,connurl,1,3
		end if
	end if
	response.redirect "url.asp"
end sub

sub move()
	dim urlid,group
	group=request("group")
	urlid=request("acter")
	if request("chkall")=1 then
		set rs=server.createobject("adodb.recordset")
		sql="update url set grouptype='"&group&"' where username='"&membername&"'"
		connurl.execute(sql)
	else
		if urlid="" then
			response.write "<script>alert('请选择您要操作的收藏地址！');history.go(-1);</script>"
			response.end
		else
			set rs=server.createobject("adodb.recordset")
			sql="update url set grouptype='"&group&"' where username='"&membername&"' and id in ("&urlid&")"
			connurl.execute(sql)
		end if
	end if
	response.redirect "url.asp"
end sub


%>