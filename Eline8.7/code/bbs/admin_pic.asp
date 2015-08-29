<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	dim admin_flag
	admin_flag="42"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
			call saveconst()
		else
			call consted()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub consted()
dim sel
%>
<form method="POST" action=admin_pic.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。<br>4、以下图片均保存于论坛pic目录中，如要更换也请将图放于该目录
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 align=left>当前使用主模板（可将设置保存到下列模板中）</th>
</tr>
<tr>
<td colspan=3 class=forumrow>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from config"
rs.open sql,conn,1,1
do while not rs.eof
if request("skinid")="" then
	if request("boardid")="" then
	if rs("active")=1 then
	sel="checked"
	skinid=rs("id")
	Forum_pic=split(rs("Forum_pic"),",")
	uForum_Pic=ubound(Forum_Pic)
	if uForum_Pic<16 then
		redim preserve Forum_Pic(16)
		Forum_Pic(16)="ao2.gif"
	end if
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_ubb=split(rs("Forum_ubb"),",")
	else
	sel=""
	end if
	else
	sel=""
	end if
else
	if rs("id")=cint(request("skinid")) then
	sel="checked"
	skinid=rs("id")
	Forum_pic=split(rs("Forum_pic"),",")
	uForum_Pic=ubound(Forum_Pic)
	if uForum_Pic<16 then
		redim preserve Forum_Pic(16)
		Forum_Pic(16)="ao2.gif"
	end if
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_ubb=split(rs("Forum_ubb"),",")
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_pic.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=forumrow><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
<select size=1>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from board where sid="&skinid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<option>没有论坛使用该模板</option>"
else
do while not rs.eof
response.write "<option>" & rs("boardtype") & "</option>"
rs.movenext
loop
end if
rs.close
set rs=nothing
%></select>
</td></tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>在线图片设置</th>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛管理员</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(0)" size="20" value="<%=Forum_pic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>超级版主</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(16)" size="20" value="<%=Forum_pic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛版主</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(1)" size="20" value="<%=Forum_pic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛贵宾</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(2)" size="20" value="<%=Forum_pic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>普通会员</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(3)" size="20" value="<%=Forum_pic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>客人或隐身会员</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(4)" size="20" value="<%=Forum_pic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>突出显示自己的颜色</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(5)" size="20" value="<%=Forum_pic(5)%>">&nbsp;&nbsp;<font color=<%=Forum_pic(5)%>>自己</font>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛图例</th>
</tr>
<tr> 
<td width="40%" class=forumrow>常规论坛－－有新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(7)" size="20" value="<%=Forum_pic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>常规论坛－－无新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(6)" size="20" value="<%=Forum_pic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>认证论坛－－有新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(15)" size="20" value="<%=Forum_pic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>认证论坛－－无新帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(14)" size="20" value="<%=Forum_pic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>联盟论坛</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(0)" size="20" value="<%=Forum_boardpic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>版面列表版面说明前面的图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(8)" size="20" value="<%=Forum_pic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>版面列表中今日发帖数图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(9)" size="20" value="<%=Forum_pic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>版面列表中总发帖数图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(10)" size="20" value="<%=Forum_pic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>版面列表中总主题数图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(11)" size="20" value="<%=Forum_pic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>论坛导航图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(12)" size="20" value="<%=Forum_pic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖子列表上方版主列表图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(13)" size="20" value="<%=Forum_pic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(13)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛主体图标</th>
</tr>
<tr> 
<td width="40%" class=forumrow>发表帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(1)" size="20" value="<%=Forum_boardpic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发表投票</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(2)" size="20" value="<%=Forum_boardpic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>回复帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(4)" size="20" value="<%=Forum_boardpic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>小字报</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(3)" size="20" value="<%=Forum_boardpic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帮助</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(5)" size="20" value="<%=Forum_boardpic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>鲜花</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(6)" size="20" value="<%=Forum_boardpic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>鸡蛋</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(7)" size="20" value="<%=Forum_boardpic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>没有新短信</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(8)" size="20" value="<%=Forum_boardpic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>有新的短消息</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(9)" size="20" value="<%=Forum_boardpic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>我发表的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(10)" size="20" value="<%=Forum_boardpic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>树形浏览帖子模式</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(11)" size="20" value="<%=Forum_boardpic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>平板形浏览帖子模式</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(12)" size="20" value="<%=Forum_boardpic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>下一篇帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(13)" size="20" value="<%=Forum_boardpic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>上一篇帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(14)" size="20" value="<%=Forum_boardpic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>刷新</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(7)" size="20" value="<%=Forum_statePic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>展开（论坛、帖子列表）</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(15)" size="20" value="<%=Forum_boardpic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>关闭展开（论坛、帖子列表）</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(16)" size="20" value="<%=Forum_boardpic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>版面列表中最后发帖</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(5)" size="20" value="<%=Forum_statePic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖子列表中分页图标</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(6)" size="20" value="<%=Forum_statePic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>短信提示音文件名</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(8)" size="20" value="<%=Forum_statePic(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>打开的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(0)" size="20" value="<%=Forum_statePic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>热门的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(1)" size="20" value="<%=Forum_statePic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>锁定的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(2)" size="20" value="<%=Forum_statePic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>固顶的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(3)" size="20" value="<%=Forum_statePic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>总固顶的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(9)" size="20" value="<%=Forum_statePic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>精华的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(4)" size="20" value="<%=Forum_statePic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>投票的主题</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(12)" size="20" value="<%=Forum_statePic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(12)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>论坛帖子图标</th>
</tr>
<tr> 
<td width="40%" class=forumrow>保存当页为文件</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_save" size="20" value="<%=Forum_TopicPic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>报告本帖给版主</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_report" size="20" value="<%=Forum_TopicPic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>显示可打印的版本</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_print" size="20" value="<%=Forum_TopicPic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖加入论坛收藏</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_bbsfav" size="20" value="<%=Forum_TopicPic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖打包邮递</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPost" size="20" value="<%=Forum_TopicPic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发送本页面给朋友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPage" size="20" value="<%=Forum_TopicPic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>把本帖加入IE收藏</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_iefav" size="20" value="<%=Forum_TopicPic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发送短信给好友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_sms" size="20" value="<%=Forum_TopicPic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>加为好友</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_friend" size="20" value="<%=Forum_TopicPic(21)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(21)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>浏览用户信息</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_userinfo" size="20" value="<%=Forum_TopicPic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>搜索用户发表帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_search" size="20" value="<%=Forum_TopicPic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户邮箱</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_Email" size="20" value="<%=Forum_TopicPic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户OICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_oicq" size="20" value="<%=Forum_TopicPic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户ICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_icq" size="20" value="<%=Forum_TopicPic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户MSN</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_msn" size="20" value="<%=Forum_TopicPic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户主页</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_homepage" size="20" value="<%=Forum_TopicPic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>引用帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_quote" size="20" value="<%=Forum_TopicPic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>编辑帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_edit" size="20" value="<%=Forum_TopicPic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>快速回复</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_freply" size="20" value="<%=Forum_TopicPic(22)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(22)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>删除帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_delete" size="20" value="<%=Forum_TopicPic(17)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(17)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>复制帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_copy" size="20" value="<%=Forum_TopicPic(18)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(18)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>加入精华帖子</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_isbest" size="20" value="<%=Forum_TopicPic(19)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(19)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>用户IP</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_ip" size="20" value="<%=Forum_TopicPic(20)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(20)%>>
</td>
</tr>
<tr > 
<th height="23" colspan="2" align=left>论坛UBB标签图标</th>
</tr>
<tr> 
<td width="40%" class=forumrow>粗体</td>
<td width="60%" class=forumrow> 
<input type="text" name="bold" size="20" value="<%=Forum_ubb(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>斜体</td>
<td width="60%" class=forumrow> 
<input type="text" name="italicize" size="20" value="<%=Forum_ubb(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>下划线</td>
<td width="60%" class=forumrow> 
<input type="text" name="underline" size="20" value="<%=Forum_ubb(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>居中</td>
<td width="60%" class=forumrow> 
<input type="text" name="center" size="20" value="<%=Forum_ubb(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>URL连接</td>
<td width="60%" class=forumrow> 
<input type="text" name="url1" size="20" value="<%=Forum_ubb(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>Email地址</td>
<td width="60%" class=forumrow> 
<input type="text" name="email1" size="20" value="<%=Forum_ubb(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖图</td>
<td width="60%" class=forumrow> 
<input type="text" name="image" size="20" value="<%=Forum_ubb(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖Flash</td>
<td width="60%" class=forumrow> 
<input type="text" name="swf" size="20" value="<%=Forum_ubb(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖ShockWave</td>
<td width="60%" class=forumrow> 
<input type="text" name="Shockwave" size="20" value="<%=Forum_ubb(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖RM影片</td>
<td width="60%" class=forumrow> 
<input type="text" name="rm" size="20" value="<%=Forum_ubb(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖Media Player影片</td>
<td width="60%" class=forumrow> 
<input type="text" name="mp" size="20" value="<%=Forum_ubb(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>帖QuickTime影片</td>
<td width="60%" class=forumrow> 
<input type="text" name="qt" size="20" value="<%=Forum_ubb(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>引用文字</td>
<td width="60%" class=forumrow> 
<input type="text" name="quote1" size="20" value="<%=Forum_ubb(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>飞行字</td>
<td width="60%" class=forumrow> 
<input type="text" name="fly" size="20" value="<%=Forum_ubb(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>移动字</td>
<td width="60%" class=forumrow> 
<input type="text" name="move" size="20" value="<%=Forum_ubb(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>发光字</td>
<td width="60%" class=forumrow> 
<input type="text" name="glow" size="20" value="<%=Forum_ubb(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>阴影字</td>
<td width="60%" class=forumrow> 
<input type="text" name="shadow" size="20" value="<%=Forum_ubb(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>心情列表</td>
<td width="60%" class=forumrow> 
<input type="text" name="smile" size="20" value="<%=Forum_ubb(17)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(17)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>背景音乐</td>
<td width="60%" class=forumrow> 
<input type="text" name="bsound" size="20" value="<%=Forum_ubb(18)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(18)%>>
</td>
</tr>
<tr> 
<td colspan=2 align=center class=forumrow><input type="submit" name="Submit" value="提 交"></td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_pic=request("Forum_pic(0)") & "," & request("Forum_pic(1)") & "," & request("Forum_pic(2)") & "," & request("Forum_pic(3)") & "," & request("Forum_pic(4)") & "," & request("Forum_pic(5)") & "," & request("Forum_pic(6)") & "," & request("Forum_pic(7)") & "," & request("Forum_pic(8)") & "," & request("Forum_pic(9)") & "," & request("Forum_pic(10)") & "," & request("Forum_pic(11)") & "," & request("Forum_pic(12)") & "," & request("Forum_pic(13)") & "," & request("Forum_pic(14)") & "," & request("Forum_pic(15)") & "," & request("Forum_pic(16)")

Forum_boardpic=request("Forum_boardpic(0)") & "," & request("Forum_boardpic(1)") & "," & request("Forum_boardpic(2)") & "," & request("Forum_boardpic(3)") & "," & request("Forum_boardpic(4)") & "," & request("Forum_boardpic(5)") & "," & request("Forum_boardpic(6)") & "," & request("Forum_boardpic(7)") & "," & request("Forum_boardpic(8)") & "," & request("Forum_boardpic(9)") & "," & request("Forum_boardpic(10)") & "," & request("Forum_boardpic(11)") & "," & request("Forum_boardpic(12)") & "," & request("Forum_boardpic(13)") & "," & request("Forum_boardpic(14)") & "," & request("Forum_boardpic(15)") & "," & request("Forum_boardpic(16)")

Forum_TopicPic=request("P_save") & "," & request("P_report") & "," & request("P_print") & "," & request("P_EmailPost") & "," & request("P_bbsfav") & "," & request("P_EmailPage") & "," & request("P_iefav") & "," & request("P_sms") & "," & request("P_search") & "," & request("P_userinfo") & "," & request("P_Email") & "," & request("P_oicq") & "," & request("P_icq") & "," & request("P_msn") & "," & request("P_homepage") & "," & request("P_quote") & "," & request("P_edit") & "," & request("P_delete") & "," & request("P_copy") & "," & request("P_isbest") & "," & request("P_ip") & "," & request("P_friend") & "," & request("P_freply")

Forum_statePic=request("Forum_statePic(0)") & "," & request("Forum_statePic(1)") & "," & request("Forum_statePic(2)") & "," & request("Forum_statePic(3)") & "," & request("Forum_statePic(4)") & "," & request("Forum_statePic(5)") & "," & request("Forum_statePic(6)") & "," & request("Forum_statePic(7)") & "," & request("Forum_statePic(8)") & "," & request("Forum_statePic(9)") & "," & request("Forum_statePic(10)") & "," & request("Forum_statePic(11)") & "," & request("Forum_statePic(12)")

Forum_ubb=request("bold") & "," & request("italicize") & "," & request("underline") & "," & request("center") & "," & request("url1") & "," & request("email1") & "," & request("image") & "," & request("swf") & "," & request("Shockwave") & "," & request("rm") & "," & request("mp") & "," & request("qt") & "," & request("quote1") & "," & request("fly") & "," & request("move") & "," & request("glow") & "," & request("shadow") & "," & request("smile") & "," & request("bsound")
'response.write Forum_Topicpic & "<br>" & Forum_boardpic & "<br>" & Forum_statePic
if request("skinid")<>"" then
sql = "update config set Forum_pic='"&Forum_pic&"',Forum_boardpic='"&Forum_boardpic&"',Forum_TopicPic='"&Forum_TopicPic&"',Forum_statePic='"&Forum_statePic&"',Forum_ubb='"&Forum_ubb&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "设置成功。<a href="&Request.ServerVariables("HTTP_REFERER")&" >点击返回</a>"
end if
end sub
%>