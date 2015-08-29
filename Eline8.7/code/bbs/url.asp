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
	call main()
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
		userid=clng(rs("id"))
		response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1 align=center><tr><td width=200 class=tablebody1 valign=top>"
		response.write "<table cellpadding=3 align=center cellspacing=1 class=tableborder1><tr><th>管理区"
		response.write "<tr><td class=tablebody2><b><img src=pic/ifolder.gif>组管理</b>&nbsp;&nbsp;<a href=urlgroup.asp?act=add><font color=red>添加</a></font>"
		response.write "<tr><td class=tablebody1>"
		for i=0 to rs("groupnum")-1
			response.write "&nbsp;&nbsp;<b>"&i+1&"、</b>"&trim(split(group,"|")(i))&"&nbsp;&nbsp;&nbsp;<a href=urlgroup.asp?act=edit&name="&trim(split(group,"|")(i))&"><font color=red>改名</font></a>&nbsp;<a href=urlgroup.asp?act=del&name="&trim(split(group,"|")(i))&"><font color=red>删除</font></a><br>"
		next
		
		response.write "</select><tr><td cellpadding=3 class=tablebody2><b><a href=urlgroup.asp?act=out><font color=red><img src=pic/delete.gif border=0>注销收藏夹</font></a></b>"
		response.write "<tr><td cellpadding=3 class=tablebody1>注：删除组将同时删除组里的所有收藏,注销收藏夹将删除个人收藏夹,并同时删除收藏的所有地址，无法恢复！</table>"
		response.write "<form name=form1 method=post action=urlact.asp><td valign=top class=tablebody1 align=left><img src=pic/fav_add1.gif><b>我的在线收藏夹</b>[全部]"
		response.write "<table width=100% cellpadding=3 cellspacing=1 align=center class=tableborder1>"
		dim brs
		for i=0 to rs("groupnum")-1
			response.write "<tr><td class=tablebody2 colspan=3><b>["&split(group,"|")(i)&"]</b>"
			set brs=server.createobject("adodb.recordset")
			sql="select * from [url] where username='"&membername&"' and grouptype='"&split(group,"|")(i)&"'"
			brs.open sql,connurl,1,1
			if brs.eof and brs.bof then
				response.write "<tr><td colspan=3 class=tablebody1>&nbsp;&nbsp;&nbsp;&nbsp;尚未收藏任何地址！"
			else
				brs.movefirst
				do while not brs.eof
					response.write "<tr><td class=tablebody2 align=center><a href=urlact.asp?act=edit&urlid="&brs("id")&"><font color=blue>"&brs("webname")&"：</font></a><td class=tablebody1><a href="&brs("weburl")&" target=_blank><img src=pic/url.gif border=0>"&brs("weburl")&"</a>"
					if brs("webtext")<>"" then response.write "&nbsp;&nbsp;(说明："&brs("webtext")&")"
					response.write "<td class=tablebody2 align=center><input type=checkbox name=acter value="&brs("id")&">"
					brs.movenext
				loop
			end if
			brs.close
			set brs=nothing
		next
		response.write "<tr><td class=tablebody2 colspan=2><b><img src=pic/url.gif><font color=red>地址管理</font></b><td class=tablebody2><input type=checkbox name=chkall value=1>全选"
		response.write "<tr><td cellpadding=3 cellspacing=1 class=tablebody1 colspan=3 valign=middle>&nbsp;&nbsp;<b><a href=urlact.asp?act=add><font color=blue>添加</font></a></b>"
		response.write "&nbsp;&nbsp;<input type=radio name=act value=del><b>删除</b>"
		response.write "&nbsp;&nbsp;<input type=radio name=act value=move><b>移动：</b>到组&nbsp;<select name=group>"
		for i=0 to rs("groupnum")-1
			response.write "<option value="&trim(split(group,"|")(i))&">"&trim(split(group,"|")(i))&"</option>"
		next
		response.write "</select>&nbsp;&nbsp;<b>编辑: </b>点击收藏名进行相关的编辑！<tr><td class=tablebody2 colspan=3 align=center><input type=submit value=确定>"
		response.write "</table></form></table>"
	end if
end sub
%>