<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="25"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
if request("action")="save" then
call savegroup()
elseif request("action")="savedit" then
call savedit()
elseif request("action")="del" then
call del()
elseif request("action")="group" then
call gradeinfo()
elseif request("action")="addgroup" then
call addgroup()
elseif request("action")="editgroup" then
call editgroup()
elseif request("action")="delgroup" then
call delgroup()
else
call usergroup()
end if
end sub

sub usergroup()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="4" >用户组管理&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF>[添加用户组]</font></a></th>
</tr>
<tr align=center>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>用户组</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>用户数量</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>编辑</B></td>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>列出用户</B></td>
</tr>
<%
dim trs
set rs=conn.execute("select * from usergroups order by usergroupID")
do while not rs.eof
set trs=conn.execute("select count(*) from [user] where UserGroupID="&rs("UserGroupID"))
%>
<tr align=center>
<td height="23" width="30%" class="Forumrow"><%=rs("title")%></td>
<td height="23" width="20%" class="Forumrow"><%=trs(0)%></td>
<td height="23" width="20%" class="Forumrow"><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a><%if rs("UserGroupID")>8 then%> | <a href="?action=delgroup&groupid=<%=rs("UserGroupID")%>">删除</a><%end if%></td>
<td height="23" width="30%" class="Forumrow"><a href="admin_user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出所有用户</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr><td colspan=4 height=25 class="Forumrow"><B>说明</B>：<BR>①在这里您可以设置各个用户组在论坛中的默认权限，论坛默认用户组不能删除和编辑用户组名<BR>②可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作，对某个组用户进行管理请点击该组用户列表<BR>③可以删除和编辑新添加的用户组<BR>④添加组后最好到等级中添加相关组的等级，具体说明请看添加等级页面，如果不添加相应等级，则该组用户的等级为用户组名，等级图象为论坛最低等级图象<BR>⑤<B>如果删除用户组，则该用户组所包含的用户将自动转到注册用户组，同时删除在等级中和该用户组关联的等级，并更新该用户组所包含的用户的等级为注册用户组按照文章计算的等级</B></td></tr>
</table>
<%
end sub

sub delgroup()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
'更新用户等级数据
Server.ScriptTimeout=999999
dim UserGrade
set rs=conn.execute("select userid,article from [user] where UserGroupID="&request("groupid"))
do while not rs.eof
	set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where UserGroupID=4 and Minarticle<="&rs(1)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
	conn.execute("update [user] set userclass='"&rs(0)&"',titlepic='"&rs(1)&"' where userid="&rs(0))
rs.movenext
loop
set rs=nothing
'更新用户组数据
conn.execute("update [user] set UserGroupID=4 where UserGroupID="&request("groupid"))
'删除关联等级
conn.execute("delete from usertitle where UserGroupID="&request("groupid"))
'删除用户组
conn.execute("delete from usergroups where UserGroupID="&request("groupid"))
response.write "删除成功！"
end sub

sub editgroup()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "请输入用户组名称！"
	exit sub
	end if
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & request.form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from usergroups where usergroupid="&request("groupid")
	rs.open sql,conn,1,3
	rs("title")=request.form("title")
	rs("GroupSetting")=GroupSetting
	rs.update
	rs.close
	set rs=nothing
	response.write "修改成功！"
else
Dim reGroupSetting
set rs=conn.execute("select * from usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到该用户组！"
exit sub
end if
reGroupSetting=split(rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup">
<input type=hidden name="groupid" value="<%=request("groupid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><td colspan=2 height=25 class="Forumrow"><B>说明</B>：<BR>①在这里您可以设置各个用户组在论坛中的默认权限，论坛默认用户组不能删除和编辑用户组名<BR>②可以删除和编辑新添加的用户组</td></tr>
<tr> 
<th height="23" colspan="2">编辑用户组：<%=rs("title")%></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>用户组名称</td>
<td height="23" width="40%" class=Forumrow><input size=35 name="title" type=text value="<%=rs("title")%>"  <%if Cint(request("GroupID"))<9 then%>disabled<%end if%>></td>
</tr>
<%if Cint(request("GroupID"))<9 then%>
<input name="title" type=hidden value="<%=rs("title")%>">
<%end if%>
<tr> 
<th height="23" colspan="2"  align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=Forumrow>是<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;否<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2"  align=left>＝＝发帖权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布新主题</td>
<td height="23" width="40%" class=Forumrow>是<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;否<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以回复自己的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;否<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以上传附件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;否<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>>
&nbsp;发帖可以上传<input name="canupload" type=radio value="2" <%if reGroupSetting(7)=2 then%>checked<%end if%>>&nbsp;回复可以上传<input name="canupload" type=radio value="3" <%if reGroupSetting(7)=3 then%>checked<%end if%>>
</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一次最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一天最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="ba2" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布新投票</td>
<td height="23" width="40%" class=Forumrow>是<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;否<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以参与投票</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;否<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布小字报</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;否<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2"  align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;否<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除自己的帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;否<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝其他权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;否<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;否<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛事件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;否<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可浏览论坛展区的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba1" type=radio value="1"  <%if reGroupSetting(49)=1 then%>checked<%end if%>>&nbsp;否<input name="ba1" type=radio value="0" <%if reGroupSetting(49)=0 then%>checked<%end if%>></td>
</tr>
<input type=hidden value=0 name="ba4">
<input type=hidden value=0 name="ba5">
<input type=hidden value=0 name="ba6">
<input type=hidden value=0 name="ba7">
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;否<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;否<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;否<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;否<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以进行帖子总固顶操作
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;否<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发帖用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;否<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;否<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;否<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;否<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;否<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;否<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;否<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;否<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理用户权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;否<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;否<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有审核帖子的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;否<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有进入隐含论坛的权限<BR>在用户组中开放则可进入所有隐含论坛<BR>在论坛权限管理中设置可进入设置的版面<BR>在用户权限管理中设置可进入设置的版面
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;否<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有论坛文件管理权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminfile" type=radio value="1" <%if reGroupSetting(48)=1 then%>checked<%end if%>>&nbsp;否<input name="canadminfile" type=radio value="0" <%if reGroupSetting(48)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有帖子主题颜色权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba3" type=radio value="1" <%if reGroupSetting(51)=1 then%>checked<%end if%>>&nbsp;否<input name="ba3" type=radio value="0" <%if reGroupSetting(51)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;否<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多发送用户
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>信箱大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr bgcolor=<%=Forum_body(2)%>> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow><input type="submit" name="submit" value="提 交"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub addgroup()
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "请输入用户组名称！"
	exit sub
	end if
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & request.form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	'response.write GroupSetting
	'response.end
	set rs=conn.execute("select * from usertitle where usertitle='"&request.form("title")&"'")
	if not (rs.eof and rs.bof) then
		Errmsg="<br><li>该等级名称已经存在。"
		call dvbbs_error()
		exit sub
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from usergroups where title='"&request.form("title")&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("title")=request.form("title")
		rs("GroupSetting")=GroupSetting
		rs.update
	else
		Errmsg="<br><li>该用户组名称已经存在。"
		call dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	dim maxid,maxuserclass,toptitlepic
	set rs=conn.execute("select top 1 * from usergroups order by usergroupid desc")
	maxid=rs("usergroupid")
	set rs=conn.execute("select max(userclass) from usertitle")
	maxuserclass=rs(0)+1
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from usertitle where usertitle='"&request.form("title")&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("userclass")=maxuserclass
		rs("usertitle")=request.form("title")
		rs("minarticle")=-1
		rs("titlepic")="level10.gif"
		rs("usergroupid")=maxid
		rs.update
	else
		Errmsg="<br><li>该等级名称已经存在。"
		call dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	response.write "添加成功，其中同时对该组新增了论坛等级，新等级采用了默认的图片设置，您可以到等级管理中进行修改！"
else
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=addgroup">
<tr><td colspan=2 height=25 class="Forumrow"><B>说明</B>：<BR>①可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作，对某个组用户进行管理请点击该组用户列表<BR>②可以删除和编辑新添加的用户组<BR>③<B>添加组后最好到等级中添加相关组的等级</B>，具体说明请看添加等级页面，如果不添加相应等级，则该组用户的等级为用户组名，等级图象为论坛最低等级图象</td></tr>
<tr> 
<th height="23" colspan="2" >添加新的用户组</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>用户组名称</td>
<td height="23" width="40%" class=Forumrow><input size=35 name="title" type=text></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=Forumrow>是<input name="canview" type=radio value="1" checked>&nbsp;否<input name="canview" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewuserinfo" type=radio value="1" checked>&nbsp;否<input name="canviewuserinfo" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewpost" type=radio value="1" checked>&nbsp;否<input name="canviewpost" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewbest" type=radio value="1" checked>&nbsp;否<input name="canviewbest" type=radio value="0"></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>发帖权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布新主题</td>
<td height="23" width="40%" class=Forumrow>是<input name="cannewpost" type=radio value="1" checked>&nbsp;否<input name="cannewpost" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以回复自己的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canreplymytopic" type=radio value="1" checked>&nbsp;否<input name="canreplymytopic" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canreplytopic" type=radio value="1" checked>&nbsp;否<input name="canreplytopic" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canpostagree" type=radio value="1" checked>&nbsp;否<input name="canpostagree" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="5"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以上传附件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canupload" type=radio value="1" checked>&nbsp;否<input name="canupload" type=radio value="0">&nbsp;发帖可以上传<input name="canupload" type=radio value="2">&nbsp;回复可以上传<input name="canupload" type=radio value="3"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一次最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="3"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一天最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="ba2" type=text size=4 value="30"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="100"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布新投票</td>
<td height="23" width="40%" class=Forumrow>是<input name="canpostvote" type=radio value="1" checked>&nbsp;否<input name="canpostvote" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以参与投票</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvote" type=radio value="1" checked>&nbsp;否<input name="canvote" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布小字报</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansmallpaper" type=radio value="1" checked>&nbsp;否<input name="cansmallpaper" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="50" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝帖子/主题编辑权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="caneditmytopic" type=radio value="1">&nbsp;否<input name="caneditmytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除自己的帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candelmytopic" type=radio value="1">&nbsp;否<input name="candelmytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovemytopic" type=radio value="1">&nbsp;否<input name="canmovemytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosemytopic" type=radio value="1">&nbsp;否<input name="canclosemytopic" type=radio value="0" checked></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝其他权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansearch" type=radio value="1" checked>&nbsp;否<input name="cansearch" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmailtopic" type=radio value="1">&nbsp;否<input name="canmailtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodify" type=radio value="1" checked>&nbsp;否<input name="canmodify" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛事件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canvieweven" type=radio value="1" checked>&nbsp;否<input name="canvieweven" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可浏览论坛展区的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba1" type=radio value="1">&nbsp;否<input name="ba1" type=radio value="0" checked></td>
</tr>
<input type=hidden value=0 name="ba4">
<input type=hidden value=0 name="ba5">
<input type=hidden value=0 name="ba6">
<input type=hidden value=0 name="ba7">
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="candeltopic" type=radio value="1">&nbsp;否<input name="candeltopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmovetopic" type=radio value="1">&nbsp;否<input name="canmovetopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canclosetopic" type=radio value="1">&nbsp;否<input name="canclosetopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cantoptopic" type=radio value="1">&nbsp;否<input name="cantoptopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以进行帖子总固顶操作
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusesign" type=radio value="1">&nbsp;否<input name="canusesign" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发帖用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canawardtopic" type=radio value="1">&nbsp;否<input name="canawardtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canaward" type=radio value="1">&nbsp;否<input name="canaward" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canmodifytopic" type=radio value="1">&nbsp;否<input name="canmodifytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbesttopic" type=radio value="1">&nbsp;否<input name="canbesttopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAnnounce" type=radio value="1">&nbsp;否<input name="canAnnounce" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminAnnounce" type=radio value="1">&nbsp;否<input name="canAdminAnnounce" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminPaper" type=radio value="1">&nbsp;否<input name="canAdminPaper" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canAdminUser" type=radio value="1">&nbsp;否<input name="canAdminUser" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canDelUserTopic" type=radio value="1">&nbsp;否<input name="canDelUserTopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canviewip" type=radio value="1">&nbsp;否<input name="canviewip" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminip" type=radio value="1">&nbsp;否<input name="canadminip" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理用户权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="adminpermission" type=radio value="1">&nbsp;否<input name="adminpermission" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canbatchtopic" type=radio value="1">&nbsp;否<input name="canbatchtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有审核帖子的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canusetitle" type=radio value="1" checked>&nbsp;否<input name="canusetitle" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有进入隐含论坛的权限<BR>在用户组中开放则可进入所有隐含论坛<BR>在论坛权限管理中设置可进入设置的版面<BR>在用户权限管理中设置可进入设置的版面
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canuseface" type=radio value="1" checked>&nbsp;否<input name="canuseface" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有论坛文件管理权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="canadminfile" type=radio value="1">&nbsp;否<input name="canadminfile" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>有帖子主题颜色权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="ba3" type=radio value="1">&nbsp;否<input name="ba3" type=radio value="0" checked></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="cansendsms" type=radio value="1" checked>&nbsp;否<input name="cansendsms" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>最多发送用户
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="5"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="300"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>信箱大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="100"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow><input type="submit" name="submit" value="提 交"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub gradeinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center bordercolor=<%=Forum_body(1)%>>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>用户门派管理</b>：您可以添加修改或者删除论坛门派。</font></td>
</tr>
<%if request("action")="edit" then%>
<form method="POST" action=admin_group.asp?action=savedit>
<%
	set rs=conn.execute("select * from GroupName where id="&request("id"))
%>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>修改门派</b> | <a href=admin_group.asp><font color="<%=Forum_body(7)%>">添加门派</font></a></font></td>
</tr>
<tr> 
<td width="30%"><font color="<%=Forum_body(7)%>"><b>门派</B>名称</font></td>
<td width="70%"> 
<input type="text" name="Groupname" size="35" value="<%=rs("Groupname")%>">&nbsp;<input type="submit" name="Submit" value="提 交">
<input type=hidden name=id value="<%=request("id")%>">
</td>
</tr>
<%set rs=nothing%>
<%else%>
<form method="POST" action=admin_group.asp?action=save>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>添加门派</b></font></td>
</tr>
<tr> 
<td width="30%"><font color="<%=Forum_body(7)%>"><b>门派</B>名称</font></td>
<td width="70%"> 
<input type="text" name="Groupname" size="35">&nbsp;<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</form>
<%end if%>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>管理门派</b></font></td>
</tr>
<%
	set rs=conn.execute("select * from GroupName")
	do while not rs.eof
%>
<tr> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>">
<a href="admin_group.asp?id=<%=rs("id")%>&action=edit"><font color="<%=Forum_body(7)%>">修改</font></a> | <a href="admin_group.asp?id=<%=rs("id")%>&action=del"><font color="<%=Forum_body(7)%>">删除</font></a> | <%=rs("GroupName")%></font>
</td>
</tr>
<%
	rs.movenext
	loop
	set rs=nothing
%>
</table>
<%
end sub

sub savegroup()
	conn.execute("insert into GroupName (GroupName) values ('"&request("GroupName")&"')")
%>
<center><p><b>添加成功！</b>
<%
end sub
sub savedit()
	conn.execute("update GroupName set GroupName='"&request("GroupName")&"' where id="&request("id"))
%>
<center><p><b>修改成功！</b>
<%
end sub
sub del()
	conn.execute("delete from GroupName where id="&request("id"))
%>
<center><p><b>删除成功！</b>
<%
end sub
%>