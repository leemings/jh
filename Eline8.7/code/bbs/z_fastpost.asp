<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_fastpost_conn.asp"-->
<%
dim msg
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
stats="快速发帖列表"
call nav()
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
%>
<!--#include file="z_controlpanel.asp"-->
<br>
<%
	select case request("action")
	case "info"
		call info()
	case "new"
		call addF()
	case "edit"
		call editF()
	case "newsave"
		call newsave()
	case "editsave"
		call editsave()
	case "删除"
		call DelF()
	case "清空快速发帖"
		call AllDelF()
	case else
		call info()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
connfastpost.close
set connfastpost=nothing
call footer()

sub info()%>
	<form action="z_fastpost.asp" method=post name=inbox>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
		<tr>
			<th valign=middle width="25%">标签名称</td>
			<th valign=middle width="*">标签内容</td>
			<th valign=middle width="5%">操作</td>
		</tr>
		<%set rs=server.createobject("adodb.recordset")
		sql="select * from FastPost where username='"&trim(membername)&"'"
		rs.open sql,connfastpost,1,1
		if rs.eof and rs.bof then
			%><tr>
				<td class=tablebody1 align=center valign=middle colspan=6>您的快速发帖列表中没有任何内容。</td>
			</tr>
		<%else
			do while not rs.eof%>
				<tr>
					<td align=center valign=middle class=tablebody1><a href="z_fastpost.asp?action=edit&id=<%=rs("ID")%>"><%=htmlencode(rs("LabelName"))%></a></td>
					<td align=left valign=middle class=tablebody1><blockquote><%=htmlencode(rs("labelcontent"))%></blockquote></td>
					<td align=center class=tablebody1><input type=checkbox name=id value=<%=rs("id")%>></td>
				</tr>
				<%rs.movenext
			loop
		end if
		rs.close
		set rs=nothing%>
		<tr> 
			<td align=right valign=middle colspan=3 class=tablebody2><input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=button name=action onclick="location.href='z_fastpost.asp?action=new'" value="添加快速发帖">&nbsp;<input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除">&nbsp;<input type=submit name=action onclick="{if(confirm('确定清除所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空快速发帖"></td>
		</tr>
	</table>
	</form>
<%end sub

sub delF()
	dim delid

	delid=replace(request.form("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id in ("&delid&")")(0)<1 then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		founderr=true
		exit sub
	else
		connfastpost.execute("delete from Fastpost where username='"&trim(membername)&"' and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>您已经删除选定的快速发帖。"
		call dvbbs_suc()
	end if
end sub

sub AllDelF()
	connfastpost.execute("delete from Fastpost where username='"&trim(membername)&"'")
	sucmsg=sucmsg+"<br>"+"<li>您已经删除了所有快速发帖列表。"
	call dvbbs_suc()
end sub

sub addF()%>
	<script src="inc/ubbcode.js"></script>
	<form method=POST name=frmAnnounce action="?action=newsave">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<TBODY>
		<tr><th width=100% colspan="2" height=23 id=tabletitlelink>用户快速发帖用语 | <a href=z_fastpost.asp>返回设置首页</a></th></tr>
		<tr>
			<td width="20%" valign=middle class=tablebody1 align=left><b>用户快速发帖用语说明：</b><br><br><li>您可以在发帖时直接调<br>&nbsp;&nbsp;&nbsp;用此用语来加速发帖<br><br><li>请您尽可能少用特殊帖</td>
			<td width="80%" valign=middle class=tablebody1 align=left>标签名称&nbsp;<input name=subject size=50 maxlength=20>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><strong>*</strong></font>不得超过 10 个汉字或20个英文字符<br><!--#include file="inc/getubb.asp"--><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title="请写入您常用的短语（支持各种特殊帖形式），可以使用Ctrl+Enter直接提交" class=FormClass onkeydown=ctlent()></textarea></td>
		</tr>
		<tr><td class=tablebody2 valign=middle align=center colspan="2" height="36"><input type=Submit value="提 交" name=Submit>&nbsp;&nbsp;<input type=reset name=Submit2 value="清 除"></td></tr>
	</TBODY>
	</table>
	</form>
<%end sub

sub editF()
	dim editid

	editid=replace(request("id"),"'","")
	if editid="" or isnull(editid) then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id="&editid&"")(0)<1 then
		Errmsg=Errmsg+"<li>"+"请选择相关参数。"
		founderr=true
		exit sub
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from FastPost where id="&editid&""
	rs.open sql,connfastpost,1,1
	%><script src="inc/ubbcode.js"></script>
	<form method=POST name=frmAnnounce action="?action=editsave&id=<%=rs("ID")%>">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<TBODY>
		<tr><th width=100% colspan="2" height=23 id=tabletitlelink>用户快速发帖用语 | <a href=z_fastpost.asp>返回设置首页</a></th></tr>
		<tr>
			<td width="20%" valign=middle class=tablebody1 align=left><b>用户快速发帖用语说明：</b><br><br><li>您可以在发帖时直接调<br>&nbsp;&nbsp;&nbsp;用此用语来加速发帖<br><br><li>请您尽可能少用特殊帖</td>
			<td width="80%" valign=middle class=tablebody1 align=left>标签名称&nbsp;<input name=subject size=50 maxlength=20 value=<%=rs("labelname")%>>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><strong>*</strong></font>不得超过 10 个汉字或20个英文字符<br><!--#include file="inc/getubb.asp"--><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title="请写入您常用的短语（支持各种特殊帖形式），可以使用Ctrl+Enter直接提交" class=FormClass onkeydown=ctlent()><%=rs("labelcontent")%></textarea></td>
		</tr>
		<tr><td class=tablebody2 valign=middle align=center colspan="2" height="36"><input type=Submit value="提 交" name=Submit>&nbsp;&nbsp;<input type=reset name=Submit2 value="清 除"></td></tr>
	</TBODY>
	</table>
	</form>
	<%rs.close
	set rs=nothing
end sub

sub newsave()
	dim labelname
	dim labelcontent
	
	labelname=replace(request("subject"),"'","")
	labelcontent=replace(request("content"),"'","")
	if labelname="" then
		errmsg=errmsg+"<li>您忘记填写标签名称了吧。"
		founderr=true
		exit sub
	elseif labelContent="" then
		errmsg=errmsg+"<li>您忘记填写标签内容了吧。"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and labelname='"&labelname&"'")(0)>0 then
		Errmsg=Errmsg+"<li>该标签名称已经存在，不能再次添加。"
		founderr=true
		exit sub
	end if
	connfastpost.execute("insert into fastpost (username,labelname,labelcontent) values ('"&membername&"','"&labelname&"','"&labelcontent&"')")
	sucmsg=sucmsg+"<li>恭喜您，已经成功的添加了快速发帖。"
	call dvbbs_suc()
end sub

sub editsave()
	dim editid
	dim labelname
	dim labelcontent
	
	editid=replace(request("id"),"'","")
	labelname=replace(request("subject"),"'","")
	labelcontent=replace(request("content"),"'","")
	if editid="" or isnull(editid) then
		Errmsg=Errmsg+"<li>请选择相关参数。"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id="&editid&"")(0)<1 then
		Errmsg=Errmsg+"<li>请选择相关参数。"
		founderr=true
		exit sub
	elseif labelname="" then
		errmsg=errmsg+"<li>您忘记填写标签名称了吧。"
		founderr=true
		exit sub
	elseif labelContent="" then
		errmsg=errmsg+"<li>您忘记填写标签内容了吧。"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and labelname='"&labelname&"' and id<>"&editid&"")(0)>0 then
		Errmsg=Errmsg+"<li>不能修改为已经存在的标签名称。"
		founderr=true
		exit sub
	end if
	connfastpost.execute("update fastpost set labelname='"&labelname&"',labelcontent='"&labelcontent&"' where id="&editid&"")
	sucmsg=sucmsg+"<li>恭喜您，已经成功的修改了快速发帖。"
	call dvbbs_suc()
end sub
%>