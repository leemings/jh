<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="01"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
		call saveconst()
		elseif request("action")="restore" then
		call restore()
		else
		call consted()
		end if
		conn.close
		set conn=nothing
	end if


sub consted()
dim sel
%>
<form method="POST" action="admin_setting.asp?action=save">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height=25 colspan=2 align=left><B>说明</B>：这里的设置代表了所有论坛中所有的设置，包括总设置和分论坛，在这里您可以设置几套不同的方案以供各个不同的分论坛使用。</td>
</tr>
<tr> 
<td height=25 colspan=2 align=left>
<table width=85% align=center>
<tr>
<td width=50% ><a name=top></a><a href="#setting1">[论坛设置]</a></td>
<td width=50% ><a href="#setting17">[防刷新机制]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting3">[基本信息]</a></td>
<td width=50% ><a href="#setting4">[联系资料]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting16">[帖子选项]</a></td>
<td width=50% ><a href="#setting6">[悄悄话选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting7">[论坛首页选项]</a></td>
<td width=50% ><a href="#setting8">[用户与注册选项]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting19">[门派设置]</a></td>
<td width=50% ><a href="#setting10">[系统设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting18">[论坛分页设置]</a></td>
<td width=50% ><a href="#setting12">[在线和用户来源]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting13">[邮件选项]</a></td>
<td width=50% ><a href="#setting14">[上传设置]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting15">[用户选项(签名、头衔、排行等)]</a></td>
<td width=50% ></td>
</tr>
</table>
</td>
</tr> 
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2  class="tableHeaderText" height=25>当前使用设置（可将当前设置保存到下列设置模板中，选中即可）
</th></tr>
<tr> 
<td width="100%" colspan=2 class="forumrow" height=25><input type="checkbox" name="checkall" checked value="1">&nbsp;全选：如果要使下面设置对所有模板生效，请全选，建议全选
</td>
</tr>
<tr> 
<td width="100%" colspan=2 class="forumrow">
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
	Forum_Setting=split(rs("Forum_setting"),",")
	if Forum_Setting(12)="" then Forum_Setting(12)="0"
	if Forum_Setting(13)="" then Forum_Setting(13)="0"
	if Forum_Setting(16)="" then Forum_Setting(16)="0"
	if Forum_Setting(17)="" then Forum_Setting(17)="1"
	if Forum_Setting(18)="" then Forum_Setting(18)="1"
	if Forum_Setting(27)="" then Forum_Setting(27)="0"
	if Forum_Setting(39)="" then Forum_Setting(39)=Forum_Setting(38)
	if Forum_Setting(43)="" then Forum_Setting(43)="0"
	if Forum_Setting(45)="" then Forum_Setting(45)="0"
	if Forum_Setting(58)="" then Forum_Setting(58)=Forum_Setting(57)
	if Forum_Setting(71)="" then Forum_Setting(71)="0"
	if Forum_Setting(72)="" then Forum_Setting(72)="00"
	uForum_Setting=ubound(Forum_Setting)
	if uForum_Setting<77 then
		redim preserve Forum_Setting(77)
		Forum_Setting(73)="0"
		Forum_Setting(74)="1"
		Forum_Setting(75)="1"
		Forum_Setting(76)="1"
		Forum_Setting(77)="3"
	end if
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
	Forum_Setting=split(rs("Forum_setting"),",")
	if Forum_Setting(12)="" then Forum_Setting(12)="0"
	if Forum_Setting(13)="" then Forum_Setting(13)="0"
	if Forum_Setting(16)="" then Forum_Setting(16)="0"
	if Forum_Setting(17)="" then Forum_Setting(17)="1"
	if Forum_Setting(18)="" then Forum_Setting(18)="1"
	if Forum_Setting(27)="" then Forum_Setting(27)="0"
	if Forum_Setting(39)="" then Forum_Setting(39)=Forum_Setting(38)
	if Forum_Setting(43)="" then Forum_Setting(43)="0"
	if Forum_Setting(45)="" then Forum_Setting(45)="0"
	if Forum_Setting(58)="" then Forum_Setting(58)="32"
	if Forum_Setting(71)="" then Forum_Setting(71)="0"
	if Forum_Setting(72)="" then Forum_Setting(72)="00"
	uForum_Setting=ubound(Forum_Setting)
	if uForum_Setting<77 then
		redim preserve Forum_Setting(77)
		Forum_Setting(73)="0"
		Forum_Setting(74)="1"
		Forum_Setting(75)="1"
		Forum_Setting(76)="1"
		Forum_Setting(77)="3"
	end if
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_setting.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2 class="Forumrow"><B>当前使用该设置的分论坛</B> | <a href="?action=restore&skinid=<%=skinid%>"><B>还原论坛常规设置</B></a> <BR><BR>以下分论坛将使用本设置，如需修改论坛使用设置，可到论坛管理中重新设置<BR>
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
<tr> 
<th height="25" colspan="2" align=left id=TableTitleLink><a name="setting1"></a><b>论坛设置</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>版面语言</U></td>
<td width="50%" class=Forumrow> 
<select size=1 name="forum_setting(9)">
<option value="lang_gb.asp">简体中文
</select>
</td>
</tr>
<!--shinzeal加入cookies防欺骗设置-->
<tr> 
<td width="50%" class=Forumrow><U>Cookies欺骗设置</U><BR>如果您需要防止Cookies欺骗，请开启，非固定IP用户可能无法长期保存Cookies信息。</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(72)" value=0 <%if cint(Forum_Setting(72))=0 then%>checked<%end if%>>关闭
<input type=radio name="Forum_Setting(72)" value=1 <%if cint(Forum_Setting(72))=1 then%>checked<%end if%>>对版主以上开启
<input type=radio name="Forum_Setting(72)" value=2 <%if cint(Forum_Setting(72))=2 then%>checked<%end if%>>对所有用户开启
</td>
</tr>
<!--shinzeal添加完毕-->

<tr> 
<td width="50%" class=Forumrow><U>论坛当前状态</U><BR>维护期间可设置关闭论坛使用</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(21)" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>打开&nbsp;
<input type=radio name="Forum_Setting(21)" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>关闭&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>维护说明</U><BR>在论坛关闭情况下显示，支持html语法</td>
<td width="50%" class=Forumrow> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否起用定时开关论坛</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="forum_setting(69)" value=0 <%if cint(Forum_Setting(69))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(69)" value=1 <%if cint(Forum_Setting(69))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛开放时间</U><BR>请确认您已经设置起用定时开关功能<BR>起止小时用符号“|”分开</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="forum_setting(70)" value="<%=forum_setting(70)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页模式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(4)" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>新版样式&nbsp;
<input type=radio name="Forum_Setting(4)" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>旧版样式&nbsp;
</td>
</tr>
<td width="50%" class=Forumrow> <U>导航菜单模式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(71)" value=0 <%if cint(Forum_Setting(71))=0 then%>checked<%end if%>>顶部固定&nbsp;
<input type=radio name="Forum_Setting(71)" value=1 <%if cint(Forum_Setting(71))=1 then%>checked<%end if%>>左侧滑动&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>论坛基本信息</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛名称</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛的url</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛的建站日期</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CreateDate" size="35" value="<%=Forum_info(12)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主页名称</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>主页URL</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛首页Logo地址</U><BR>显示在论坛首页，添加论坛如果没有填写logo地址，则使用该内容</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛图片目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发帖心情目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发帖表情目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="emoturl" size="35" value="<%=Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛头像目录</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="UserFaceurl" size="35" value="<%=Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>版权信息</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting4"></a><b>论坛联系资料</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛管理员Email</U><BR>给用户发送邮件时，显示的来源Email信息</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>悄悄话选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>新短消息弹出窗口</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(10)" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(10)" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting7"></a><b>论坛首页选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>首页显示论坛深度</U><BR>0代表一级，1代表2级，以此类推<BR>设置过大的论坛深度将影响论坛整体性能，请根据自己论坛情况做设置，建议设置为1</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="forum_setting(5)" value="<%=forum_setting(5)%>"> 级
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>快速登录位置</U></td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(31)">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>顶部
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>底部
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>不显示
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示过生日会员</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(29)" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(29)" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示周年纪念会员</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(43)" value=0 <%if cint(Forum_Setting(43))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(43)" value=1 <%if cint(Forum_Setting(43))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页是否显示点歌列表(停用)</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(12)" value=0 <%if cint(Forum_Setting(12))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(12)" value=1 <%if cint(Forum_Setting(12))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页是否显示回收站</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(13)" value=0 <%if cint(Forum_Setting(13))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(13)" value=1 <%if cint(Forum_Setting(13))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页是否显示明星</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(16)" value=0 <%if cint(Forum_Setting(16))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%if cint(Forum_Setting(16))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>明星第一列的显示方式</U></td>
<td width="50%" class=Forumrow><select name="Forum_Setting(17)"> 
<option value=1 <%if cint(Forum_Setting(17))=1 then%>selected<%end if%>>今日发帖量</option>
<option value=2 <%if cint(Forum_Setting(17))=2 then%>selected<%end if%>>本周发帖量</option>
<option value=3 <%if cint(Forum_Setting(17))=3 then%>selected<%end if%>>本月发帖量</option>
<option value=4 <%if cint(Forum_Setting(17))=4 then%>selected<%end if%>>今年发帖量</option>
<option value=5 <%if cint(Forum_Setting(17))=5 then%>selected<%end if%>>发帖总数</option>
<option value=6 <%if cint(Forum_Setting(17))=6 then%>selected<%end if%>>最佳男明星</option>
<option value=7 <%if cint(Forum_Setting(17))=7 then%>selected<%end if%>>最佳女明星</option>
<option value=8 <%if cint(Forum_Setting(17))=8 then%>selected<%end if%>>今日得分量</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>明星第二列的显示方式</U></td>
<td width="50%" class=Forumrow><select name="Forum_Setting(18)"> 
<option value=1 <%if cint(Forum_Setting(18))=1 then%>selected<%end if%>>今日发帖量</option>
<option value=2 <%if cint(Forum_Setting(18))=2 then%>selected<%end if%>>本周发帖量</option>
<option value=3 <%if cint(Forum_Setting(18))=3 then%>selected<%end if%>>本月发帖量</option>
<option value=4 <%if cint(Forum_Setting(18))=4 then%>selected<%end if%>>今年发帖量</option>
<option value=5 <%if cint(Forum_Setting(18))=5 then%>selected<%end if%>>发帖总数</option>
<option value=6 <%if cint(Forum_Setting(18))=6 then%>selected<%end if%>>最佳男明星</option>
<option value=7 <%if cint(Forum_Setting(18))=7 then%>selected<%end if%>>最佳女明星</option>
<option value=8 <%if cint(Forum_Setting(18))=8 then%>selected<%end if%>>今日得分量</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页是否显示“我最喜爱的论坛”</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(75)" value=0 <%if cint(Forum_Setting(75))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(75)" value=1 <%if cint(Forum_Setting(75))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>版块图片是否淡入淡出</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(76)" value=0 <%if cint(Forum_Setting(76))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(76)" value=1 <%if cint(Forum_Setting(76))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>首页Logo联盟论坛样式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(77)" value=0 <%if cint(Forum_Setting(77))=0 then%>checked<%end if%>>不淡化不滚动&nbsp;
<input type=radio name="Forum_Setting(77)" value=1 <%if cint(Forum_Setting(77))=1 then%>checked<%end if%>>只淡化不滚动&nbsp;
<input type=radio name="Forum_Setting(77)" value=2 <%if cint(Forum_Setting(77))=2 then%>checked<%end if%>>只滚动不淡化&nbsp;
<input type=radio name="Forum_Setting(77)" value=3 <%if cint(Forum_Setting(77))=3 then%>checked<%end if%>>既滚动又淡化&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>用户与注册选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否允许新用户注册</U><BR>关闭后论坛将不能注册</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(37)" value=0 <%if cint(forum_setting(37))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if cint(forum_setting(37))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>最短用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(40)" size="3" value="<%=forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>最长用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(41)" size="3" value="<%=forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>同一IP注册间隔时间</U><BR>如不想限制可填写0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(22)" size="3" value="<%=Forum_Setting(22)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email通知密码</U><BR>确认您的站点支持发送mail，所包含密码为系统随机生成</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(23)" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(23)" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>一个Email只能注册一个帐号</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(24)" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(24)" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>注册需要管理员认证</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(25)" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(25)" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送注册信息邮件</U><BR>请确认您打开了邮件功能</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(47)" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(47)" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>开启短信欢迎新注册用户</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(46)" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(46)" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>系统设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛所在时区</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>服务器时差</U></td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>脚本超时时间</U><BR>默认为300，一般不做更改</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(1)" size="3" value="<%=Forum_Setting(1)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示页面执行时间</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(30)" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(30)" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>禁止的邮件地址</U><BR>在下面指定的邮件地址将被禁止注册，每个邮件地址用“|”符号分隔<BR>本功能支持模糊搜索，如设置了eway禁止，将禁止eway@aspsky.net或者eway@dvbbs.net类似这样的注册</td>
<td width="50%" class=Forumrow> 
<input type="text" name="Forum_Setting(52)" size="50" value="<%=Forum_Setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>防刷新功能有效的页面</U><BR>请确认您打开了防刷新功能<BR>您指定的页面将有防刷新作用，用户在限定的时间内不能重复打开该页面，具有一定减少资源消耗的作用<BR>每个页面名请用“|”符号隔开</td>
<td width="50%" class=Forumrow> 
<input type="text" name="Forum_Setting(64)" size="50" value="<%=Forum_Setting(64)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>在线和用户来源</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线显示用户IP</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(28)" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>保密&nbsp;
<input type=radio name="Forum_Setting(28)" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线显示用户来源</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见<BR>开启本功能较消耗资源</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(36)" value=0 <%if cint(forum_setting(36))=0 then%>checked<%end if%>>保密&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if cint(forum_setting(36))=1 then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户当前位置</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(33)" value=0 <%if cint(forum_setting(33))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if cint(forum_setting(33))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户登录和活动时间</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(34)" value=0 <%if cint(forum_setting(34))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if cint(forum_setting(34))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户浏览器和操作系统</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(35)" value=0 <%if cint(forum_setting(35))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if cint(forum_setting(35))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示客人在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(15)" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(15)" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示用户在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(14)" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(14)" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>删除不活动用户时间</U><BR>可设置删除多少分钟内不活动用户<BR>单位：分钟，请输入数字</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(8)" size="3" value="<%=Forum_Setting(8)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>总论坛允许同时在线数</U><BR>如不想限制，可设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(26)" size="6" value="<%=Forum_Setting(26)%>">&nbsp;人
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting13"></a><b>邮件选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送邮件组件</U><BR>如果您的服务器不支持下列组件，请选择不支持</td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(2)">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>不支持 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>SMTP Server地址</U><BR>只有在论坛使用设置中打开了发送邮件功能，该填写内容方有效</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting14"></a><b>上传设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>头像上传</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(7)" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(7)" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>允许的最大头像文件大小</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(56)" size="6" value="<%=forum_setting(56)%>">&nbsp;K
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>用户选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许个人签名</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(42)" value=0 <%if forum_setting(42)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if forum_setting(42)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许用户使用头像</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(53)" value=0 <%if forum_setting(53)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if forum_setting(53)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>最大头像宽度</U><BR>定义内容为头像的最大宽度</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(57)" size="6" value="<%=forum_setting(57)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>最大头像高度</U><BR>定义内容为头像的最大高度</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(58)" size="6" value="<%=forum_setting(58)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>默认头像宽度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(38)" size="6" value="<%=forum_setting(38)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>默认头像高度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(39)" size="6" value="<%=forum_setting(39)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>使用自定义头像的最少发帖数</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=forum_setting(54)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许从其他站点上传头像</U><BR>就是是否可以直接使用http..这样的url来直接显示头像</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(55)" value=0 <%if forum_setting(55)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if forum_setting(55)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启UBB代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(65)" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(65)" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启HTML代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(66)" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(66)" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户是否开启帖图标签</U><BR>包括图片、flash、多媒体等</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(67)" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(67)" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>用户排行列表个数</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="Forum_Setting(68)" size="6" value="<%=Forum_Setting(68)%>">&nbsp;个
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户头衔</U><BR>是否允许用户自定义头衔</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(6)" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(6)" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户头衔最大长度</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(59)" size="6" value="<%=forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔最少发帖数量限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(60)" size="6" value="<%=forum_setting(60)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔注册天数限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(61)" size="6" value="<%=forum_setting(61)%>">&nbsp;天
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔上面两个条件加在一起限制</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(62)" value=0 <%if cint(Forum_Setting(62))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(62)" value=1 <%if cint(Forum_Setting(62))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔中要屏蔽的词语</U><BR>每个限制字符用“|”符号隔开</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(63)" size="50" value="<%=forum_setting(63)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>防刷新机制</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>防刷新机制</U><BR>如选择打开请填写下面的限制刷新时间<BR>对版主和管理员无效</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(19)" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(19)" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>浏览刷新时间间隔</U><BR>填写该项目请确认您打开了防刷新机制<BR>仅对帖子列表和显示帖子页面起作用</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(20)" size="3" value="<%=Forum_Setting(20)%>">&nbsp;秒
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>论坛分页设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>每页显示最多纪录</U><BR>用于论坛所有和分页有关的项目（帖子列表和浏览帖子除外）</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(11)" size="3" value="<%=Forum_Setting(11)%>">&nbsp;条
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>帖子选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>作为热门话题的最低人气值</U><BR>标准为主题回复数</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(44)" size="3" value="<%=forum_setting(44)%>">&nbsp;条
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>编辑过的帖子显示"由xxx于yyy编辑"的信息</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(48)" value=0 <%if cint(forum_setting(48))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if cint(forum_setting(48))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>管理员编辑后显示"由XXX编辑"的信息</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(49)" value=0 <%if cint(forum_setting(49))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if cint(forum_setting(49))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>等待"由XXX编辑"信息显示的时间</U><BR>允许用户编辑自己的帖子而不在帖子底部显示"由XXX编辑"信息的时限(以分钟为单位)</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(50)" size="3" value="<%=forum_setting(50)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>编辑帖子时限</U><BR>编辑处理帖子的时间限制(以分钟为单位, 1天是1440分钟) 超过这个时间限制, 只有管理员和版主才能编辑和删除帖子. 如果不想使用这项功能, 请设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(51)" size="3" value="<%=forum_setting(51)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否起用记录帖子阅读用户功能</U><BR>只有用户发言时选择记录阅读该帖用户方生效<BR>开启本功能将消耗部分系统资源<BR>发帖用户、管理员和版主可看到该帖阅读用户列表</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(3)" value=0 <%if cint(forum_setting(3))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(3)" value=1 <%if cint(forum_setting(3))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子中广告模式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(27)" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>静态广告&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>滚动新帖&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子中会员头像显示方式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(72_1)" value=0 <%if cint(left(Forum_Setting(72),1))=0 then%>checked<%end if%>>大形象&nbsp;
<input type=radio name="Forum_Setting(72_1)" value=1 <%if cint(left(Forum_Setting(72),1))=1 then%>checked<%end if%>>小形象&nbsp;
<input type=radio name="Forum_Setting(72_1)" value=2 <%if cint(left(Forum_Setting(72),1))=2 then%>checked<%end if%>>头像&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子中会员属性显示方式</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(72_2)" value=0 <%if cint(right(Forum_Setting(72),1))=0 then%>checked<%end if%>>详细&nbsp;
<input type=radio name="Forum_Setting(72_2)" value=1 <%if cint(right(Forum_Setting(72),1))=1 then%>checked<%end if%>>简单&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>帖子中是否显示存款和股票</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(45)" value=0 <%if cint(Forum_Setting(45))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(45)" value=1 <%if cint(Forum_Setting(45))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启发帖限制</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(73)" value=0 <%if cint(Forum_Setting(73))=0 then%>checked<%end if%>>不限&nbsp;
<input type=radio name="Forum_Setting(73)" value=1 <%if cint(Forum_Setting(73))=1 then%>checked<%end if%>>限制主题数&nbsp;
<input type=radio name="Forum_Setting(73)" value=2 <%if cint(Forum_Setting(73))=2 then%>checked<%end if%>>限制回复数&nbsp;
<input type=radio name="Forum_Setting(73)" value=3 <%if cint(Forum_Setting(73))=3 then%>checked<%end if%>>限制帖子数&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发帖限制比率</U><BR>每登录论坛一天可以发表的帖子个数</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(74)" size="3" value="<%=forum_setting(74)%>">&nbsp;个
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>门派设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启论坛门派</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(32)" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_Setting(32)" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>

<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
	dim Forum_copyright
	if trim(request("skinid"))="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
	else
		Forum_Setting=request.Form("Forum_Setting(0)") & "," & request.Form("Forum_Setting(1)") & "," & request.Form("Forum_Setting(2)") & "," & request.Form("Forum_Setting(3)") & "," & request.Form("Forum_Setting(4)") & "," & request.Form("Forum_Setting(5)") & "," & request.Form("Forum_Setting(6)") & "," & request.Form("Forum_Setting(7)") & "," & request.Form("Forum_Setting(8)") & "," & request.Form("Forum_Setting(9)") & "," & request.Form("Forum_Setting(10)") & "," & request.Form("Forum_Setting(11)") & "," & request.Form("Forum_Setting(12)") & "," & request.Form("Forum_Setting(13)") & "," & request.Form("Forum_Setting(14)") & "," & request.Form("Forum_Setting(15)") & "," & request.Form("Forum_Setting(16)") & "," & request.Form("Forum_Setting(17)") & "," & request.Form("Forum_Setting(18)") & "," & request.Form("Forum_Setting(19)") & "," & request.Form("Forum_Setting(20)") & "," & request.Form("Forum_Setting(21)") & "," & request.Form("Forum_Setting(22)") & "," & request.Form("Forum_Setting(23)") & "," & request.Form("Forum_Setting(24)") & "," & request.Form("Forum_Setting(25)") & "," & request.Form("Forum_Setting(26)") & "," & request.Form("Forum_Setting(27)") & "," & request.Form("Forum_Setting(28)") & "," & request.Form("Forum_Setting(29)") & "," & request.Form("Forum_Setting(30)") & "," & request.Form("Forum_Setting(31)") & "," & request.Form("Forum_Setting(32)") & "," & request.Form("Forum_Setting(33)") & "," & request.Form("Forum_Setting(34)") & "," & request.Form("Forum_Setting(35)") & "," & request.Form("Forum_Setting(36)") & "," & request.Form("Forum_Setting(37)") & "," & request.Form("Forum_Setting(38)") & "," & request.Form("Forum_Setting(39)") & "," & request.Form("Forum_Setting(40)") & "," & request.Form("Forum_Setting(41)") & "," & request.Form("Forum_Setting(42)") & "," & request.Form("Forum_Setting(43)") & "," & request.Form("Forum_Setting(44)") & "," & request.Form("Forum_Setting(45)") & "," & request.Form("Forum_Setting(46)") & "," & request.Form("Forum_Setting(47)") & "," & request.Form("Forum_Setting(48)") & "," & request.Form("Forum_Setting(49)") & "," & request.Form("Forum_Setting(50)") & "," & request.Form("Forum_Setting(51)") & "," & request.Form("Forum_Setting(52)") & "," & request.Form("Forum_Setting(53)") & "," & request.Form("Forum_Setting(54)") & "," & request.Form("Forum_Setting(55)") & "," & request.Form("Forum_Setting(56)") & "," & request.Form("Forum_Setting(57)") & "," & request.Form("Forum_Setting(58)") & "," & request.Form("Forum_Setting(59)") & "," & request.Form("Forum_Setting(60)") & "," & request.Form("Forum_Setting(61)") & "," & request.Form("Forum_Setting(62)") & "," & request.Form("Forum_Setting(63)") & "," & request.Form("Forum_Setting(64)") & "," & request.Form("Forum_Setting(65)") & "," & request.Form("Forum_Setting(66)") & "," & request.Form("Forum_Setting(67)") & "," & request.Form("Forum_Setting(68)") & "," & request.Form("Forum_Setting(69)") & "," & request.Form("Forum_Setting(70)") & "," & request.Form("Forum_Setting(71)") & "," & request.Form("Forum_Setting(72_1)") & request.Form("Forum_Setting(72_2)") & "," & request.Form("Forum_Setting(73)") & "," & request.Form("Forum_Setting(74)") & "," & request.Form("Forum_Setting(75)") & "," & request.Form("Forum_Setting(76)") & "," & request.Form("Forum_Setting(77)")
		Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT") & "," & Request("emoturl") & "," & Request("userfaceurl") & "," & Request("CreateDate")
		Forum_copyright=request("copyright")
		if request("checkall")="1" then
			sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"'"
			conn.execute(sql)
		else
			if request("skinid")<>"" then
				sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id in ( "&request("skinid")&" )"
				conn.execute(sql)
			end if
		end if
		response.write "设置成功。"
		call cache_link(request.Form("Forum_Setting(77)"))
	end if
end sub

'恢复默认设置
sub restore()
Forum_info="动网先锋论坛,http://www.dvbbs.net/,动网先锋,http://www.aspsky.net/,61.145.114.64,club@aspsky.net,images/LOGO.GIF,pic/,face/,北京时间,emot/,userface/"
Forum_Setting="0,300,1,1,0,1,1,1,40,lang_gb.asp,0,20,,,0,0,,,,0,30,0,40,0,1,0,9999,,1,0,1,1,1,1,1,1,0,1,32,32,0,10,1,,10,,1,1,1,1,10,10,0,1,0,1,500,120,,9,15,4,0,300,0,1,0,1,20,0,0|24,,"
conn.execute("update config set Forum_Setting='"&Forum_Setting&"',Forum_info='"&Forum_info&"' where id="&request("skinid"))
response.write "还原设置成功。"
end sub
%>