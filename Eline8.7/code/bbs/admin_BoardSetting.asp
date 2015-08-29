<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="md5.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or instr(session("flag"),"11")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
		call saveconst()
		else
		call consted()
		end if
		conn.close
		set conn=nothing
	end if


sub consted()
if not isnumeric(request("editid")) then
	Errmsg="<br><li>错误的版面信息"
	call dvbbs_error()
	exit sub
end if
set rs=conn.execute("select * from board where boardid="&request("editid"))
Board_Setting=split(rs("board_setting"),",")
uboard_setting=ubound(board_setting)
if uboard_setting<61 then
	redim preserve board_setting(61)
	for i=uboard_setting to 61
		board_setting(i)=0
	next
	uboard_setting=61
end if
for i=44 to 61
	if i>=50 and i<=52 then
		if board_setting(i)="" then board_setting(i)=0
	else
		if board_setting(i)="" then board_setting(i)=1
	end if
next
%>
<form method="POST" action="admin_boardsetting.asp?action=save">
<input type=hidden value="<%=request("editid")%>" name="editid">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="25" colspan="2" align=left>论坛高级设置 → <%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>本论坛作为分类论坛不允许发帖</U><BR>如果已经有帖则显示或者您可以转移到别的论坛<BR>选择了该项后所有会员均不能在本版发帖/回帖等操作</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(43)" value=0 <%if cint(Board_Setting(43))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(43)" value=1 <%if cint(Board_Setting(43))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛列表显示下属论坛风格</U><BR>当该论坛有下属论坛的时候生效</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(39)" value=0 <%if cint(Board_Setting(39))=0 then%>checked<%end if%>>列表&nbsp;
<input type=radio name="Board_Setting(39)" value=1 <%if cint(Board_Setting(39))=1 then%>checked<%end if%>>简洁&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛列表简洁风格一行版面数</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(41)" value="<%=Board_Setting(41)%>"> 个
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛风格</U><BR>该效果体现在显示帖子内容页面</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(24)" value=0 <%if cint(Board_Setting(24))=0 then%>checked<%end if%>>默认风格&nbsp;
<input type=radio name="Board_Setting(24)" value=1 <%if cint(Board_Setting(24))=1 then%>checked<%end if%>>讨论区风格&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>讨论区风格预览字数</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(25)" value="<%=Board_Setting(25)%>"> byte
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否锁定论坛</U><BR>锁定论坛只有管理员和该版面版主可进<BR></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(0)" value=0 <%if cint(Board_Setting(0))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(0)" value=1 <%if cint(Board_Setting(0))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否隐藏论坛</U><BR>隐藏论坛只有管理员和该版面版主可见和进入<BR>如果用户组或论坛权限管理或用户权限管理中允许则用户可见和进入<BR>本限制对一级论坛不生效</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(1)" value=0 <%if cint(Board_Setting(1))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(1)" value=1 <%if cint(Board_Setting(1))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否认证论坛</U><BR>加密论坛只有管理员和该版面版主可见和进入<BR>用户可进入的设定在版面管理中<BR>只有认证用户可进入</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(2)" value=0 <%if cint(Board_Setting(2))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(2)" value=1 <%if cint(Board_Setting(2))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否VIP论坛</U><BR>VIP论坛只有VIP会员和管理员可见和进入<BR>用户可以申请加入VIP会员，然后再进入</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(51)" value=0 <%if cint(Board_Setting(51))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(51)" value=1 <%if cint(Board_Setting(51))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否性别论坛</U><BR>性别论坛只有指定性别的会员和管理员可见和进入</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(52)" value=0 <%if cint(Board_Setting(52))=0 then%>checked<%end if%>>普通论坛&nbsp;
<input type=radio name="Board_Setting(52)" value=1 <%if cint(Board_Setting(52))=1 then%>checked<%end if%>>女生论坛&nbsp;
<input type=radio name="Board_Setting(52)" value=2 <%if cint(Board_Setting(52))=2 then%>checked<%end if%>>男生论坛&nbsp;
</td>
</tr>
<%'=============== 同学录 开始 ===================%>
<tr> 
<td width="50%" class=Forumrow>
<U>是否同学录论坛</U><BR>同学录论坛只有加入该班级的成员和管理员可以进入</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(50)" value=0 <%if cint(Board_Setting(50))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(50)" value=1 <%if cint(Board_Setting(50))=1 then%>checked<%end if%>>是&nbsp;&nbsp创始人&nbsp;<input type=text size=15 name="boardmaster" value="<%if not isnull(rs("txlpd")) and rs("txlpd")<>"" then%><%=split(rs("txlpd"),"$")(3)%><%end if%>"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;多创始人添加请用|分隔
<input type="hidden" name="oldboardmaster" value="<%if not isnull(rs("txlpd")) and rs("txlpd")<>"" then%><%=split(rs("txlpd"),"$")(3)%><%end if%>">
</td>
</tr>
<%'=============== 同学录 结束 ===================%>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子审核制度</U><BR>版主、管理员和开放权限用户可进行审核帖子<BR>版主、管理员和开放权限用户可直接发帖<BR>一般用户需审核后帖子方可见</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(3)" value=0 <%if cint(Board_Setting(3))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(3)" value=1 <%if cint(Board_Setting(3))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否采用版主继承制度</U><BR>如果采用该制度，则上级论坛版主可管理下级论坛相关信息</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(40)" value=0 <%if cint(Board_Setting(40))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(40)" value=1 <%if cint(Board_Setting(40))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否显示下级论坛帖子</U><BR>如果有下级论坛则显示，会加重服务器负担</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(42)" value=0 <%if cint(Board_Setting(42))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(42)" value=1 <%if cint(Board_Setting(42))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>允许同时在线数</U><BR>不限制则设置为0</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(18)" value="<%=Board_Setting(18)%>"> 人
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否起用定时开关论坛</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(21)" value=0 <%if cint(Board_Setting(21))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(21)" value=1 <%if cint(Board_Setting(21))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛开放时间</U><BR>请确认您已经设置起用定时开关功能<BR>起止小时用符号“|”分开</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(22)" value="<%=Board_Setting(22)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子列表每页记录数</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(26)" value="<%=Board_Setting(26)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>浏览帖子每页记录数</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(27)" value="<%=Board_Setting(27)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>主版主可以增删副版主</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(33)" value=0 <%if cint(Board_Setting(33))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(33)" value=1 <%if cint(Board_Setting(33))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>主版主可以修改CSS设置</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(34)" value=0 <%if cint(Board_Setting(34))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(34)" value=1 <%if cint(Board_Setting(34))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>所有版主可以修改CSS设置</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(35)" value=0 <%if cint(Board_Setting(35))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(35)" value=1 <%if cint(Board_Setting(35))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否对论坛事件中的操作者保密</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(36)" value=0 <%if cint(Board_Setting(36))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(36)" value=1 <%if cint(Board_Setting(36))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子列表默认读取数据量</U></td>
<td width="50%" class=Forumrow> 
<select size="0" name="Board_Setting(37)">
<option value="1"<%if cint(Board_Setting(37))=1 then%> selected<%end if%>>全部显示帖子</option>
<option value="2"<%if cint(Board_Setting(37))=2 then%> selected<%end if%>>五天内帖子</option>
<option value="3"<%if cint(Board_Setting(37))=3 then%> selected<%end if%>>半月内帖子</option>
<option value="4"<%if cint(Board_Setting(37))=4 then%> selected<%end if%>>一月内帖子</option>
<option value="5"<%if cint(Board_Setting(37))=5 then%> selected<%end if%>>两月内帖子</option>
<option value="6"<%if cint(Board_Setting(37))=6 then%> selected<%end if%>>四月内帖子</option>
<option value="7"<%if cint(Board_Setting(37))=7 then%> selected<%end if%>>半年内帖子</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子列表默认排序方式</U></td>
<td width="50%" class=Forumrow> 
<select size="1" name="Board_Setting(38)">
<option value="1"<%if cint(Board_Setting(38))=1 then%> selected<%end if%>>最后回复时间</option>
<option value="2"<%if cint(Board_Setting(38))=2 then%> selected<%end if%>>发帖时间</option>
<option value="3"<%if cint(Board_Setting(38))=3 then%> selected<%end if%>>标题</option>
<option value="4"<%if cint(Board_Setting(38))=4 then%> selected<%end if%>>发言人</option>
<option value="5"<%if cint(Board_Setting(38))=5 then%> selected<%end if%>>回复数</option>
<option value="6"<%if cint(Board_Setting(38))=6 then%> selected<%end if%>>浏览数</option>
</select>
</td>
</tr>
<tr> 
<th height="25" colspan="2" align=left>帖子相关设置</th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>发帖模式</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(4)" value=0 <%if cint(Board_Setting(4))=0 then%>checked<%end if%>>简单&nbsp;
<input type=radio name="Board_Setting(4)" value=1 <%if cint(Board_Setting(4))=1 then%>checked<%end if%>>高级&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>HTML代码解析</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(5)" value=0 <%if cint(Board_Setting(5))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(5)" value=1 <%if cint(Board_Setting(5))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>UBB代码解析</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(6)" value=0 <%if cint(Board_Setting(6))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(6)" value=1 <%if cint(Board_Setting(6))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖图标签</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(7)" value=0 <%if cint(Board_Setting(7))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(7)" value=1 <%if cint(Board_Setting(7))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>表情标签</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(8)" value=0 <%if cint(Board_Setting(8))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(8)" value=1 <%if cint(Board_Setting(8))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>多媒体标签</U><BR>包括Flash,RM,AVI等</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(9)" value=0 <%if cint(Board_Setting(9))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(9)" value=1 <%if cint(Board_Setting(9))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放金钱帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(10)" value=0 <%if cint(Board_Setting(10))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(10)" value=1 <%if cint(Board_Setting(10))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放积分帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(11)" value=0 <%if cint(Board_Setting(11))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(11)" value=1 <%if cint(Board_Setting(11))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放魅力帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(12)" value=0 <%if cint(Board_Setting(12))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(12)" value=1 <%if cint(Board_Setting(12))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放威望帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(13)" value=0 <%if cint(Board_Setting(13))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(13)" value=1 <%if cint(Board_Setting(13))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放文章帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(14)" value=0 <%if cint(Board_Setting(14))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(14)" value=1 <%if cint(Board_Setting(14))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放回复帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(15)" value=0 <%if cint(Board_Setting(15))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(15)" value=1 <%if cint(Board_Setting(15))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放出售帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(23)" value=0 <%if cint(Board_Setting(23))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(23)" value=1 <%if cint(Board_Setting(23))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放秘密帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(53)" value=0 <%if cint(Board_Setting(53))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(53)" value=1 <%if cint(Board_Setting(53))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放会员帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(54)" value=0 <%if cint(Board_Setting(54))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(54)" value=1 <%if cint(Board_Setting(54))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放定人帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(55)" value=0 <%if cint(Board_Setting(55))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(55)" value=1 <%if cint(Board_Setting(55))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放性别帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(56)" value=0 <%if cint(Board_Setting(56))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(56)" value=1 <%if cint(Board_Setting(56))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放高级帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(57)" value=0 <%if cint(Board_Setting(57))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(57)" value=1 <%if cint(Board_Setting(57))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放年龄帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(58)" value=0 <%if cint(Board_Setting(58))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(58)" value=1 <%if cint(Board_Setting(58))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放门派帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(59)" value=0 <%if cint(Board_Setting(59))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(59)" value=1 <%if cint(Board_Setting(59))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放显示帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(60)" value=0 <%if cint(Board_Setting(60))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(60)" value=1 <%if cint(Board_Setting(60))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放隐藏帖</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(61)" value=0 <%if cint(Board_Setting(61))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(61)" value=1 <%if cint(Board_Setting(61))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放发帖机遇功能</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(44)" value=0 <%if cint(Board_Setting(44))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(44)" value=1 <%if cint(Board_Setting(44))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>发帖机遇默认状态</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(45)" value=0 <%if cint(Board_Setting(45))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(45)" value=1 <%if cint(Board_Setting(45))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放发帖后短信回复功能</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(46)" value=0 <%if cint(Board_Setting(46))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(46)" value=1 <%if cint(Board_Setting(46))=1 then%>checked<%end if%>>开放&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>发帖后短信回复的默认状态</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(47)" value=0 <%if cint(Board_Setting(47))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(47)" value=1 <%if cint(Board_Setting(47))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放短信引起注意功能</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(48)" value=0 <%if cint(Board_Setting(48))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(48)" value=1 <%if cint(Board_Setting(48))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否开放版主评分功能</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(49)" value=0 <%if cint(Board_Setting(49))=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(49)" value=1 <%if cint(Board_Setting(49))=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>发帖后返回</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(17)" value=1 <%if cint(Board_Setting(17))=1 then%>checked<%end if%>>首页&nbsp;
<input type=radio name="Board_Setting(17)" value=2 <%if cint(Board_Setting(17))=2 then%>checked<%end if%>>论坛&nbsp;
<input type=radio name="Board_Setting(17)" value=3 <%if cint(Board_Setting(17))=3 then%>checked<%end if%>>帖子&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子内容最大字节数</U><BR>1024字节等于1K</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(16)" value="<%=Board_Setting(16)%>"> 字节
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子正文字号</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(28)" value="<%=Board_Setting(28)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>帖子正文行间距</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(29)" value="<%=Board_Setting(29)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>上传文件类型</U><BR>每种文件类型用“|”号分开</td>
<td width="50%" class=Forumrow> 
<input type=text size=50 name="Board_Setting(19)" value="<%=Board_Setting(19)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>是否起用防灌水机制</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(30)" value=0 <%if cint(Board_Setting(30))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(30)" value=1 <%if cint(Board_Setting(30))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>每次发帖间隔</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(31)" value="<%=Board_Setting(31)%>"> 秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>最多投票项目</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(32)" value="<%=Board_Setting(32)%>"> 个
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type=hidden value="0" name="Board_Setting(20)">
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
	if not isnumeric(request("editid")) then
		Errmsg=Errmsg+"<br>"+"<li>错误的版面参数"
		call dvbbs_error()
		exit sub
	else
		'============= 同学录 开始 ==================
		dim txlpd,txlpd1
		if cint(request("Board_Setting(50)"))=1 and request("BoardMaster")="" then
			Errmsg=Errmsg+"<br>"+"<li>错误的同学录创始人参数"
			call dvbbs_error()
			exit sub
		end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from board where boardid="&request("editid"),conn,1,3
		if cint(request("Board_Setting(50)"))=1 then
			if isnull(rs("txlpd")) or rs("txlpd")="" then
				txlpd1=split(now()&"$$$$$$0","$")
			else
				txlpd1=split(rs("txlpd"),"$")
			end if
		else
			txlpd1=split(now()&"$$$$$$0","$")
		end if
		if cint(request("Board_Setting(50)"))=1 then
			txlpd=txlpd1(0)&"$"&txlpd1(1)&"$"&txlpd1(2)&"$"&request("boardmaster")&"$"&txlpd1(4)&"$"&txlpd1(5)&"$1"
			rs("txlpd")=txlpd
		else
			txlpd=now()&"$$$$$$0"
			rs("txlpd")=txlpd
		end if
		dim iboard_setting,barr,boardmaster_1
		boardmaster_1=""
		if not isnull(rs("boardmaster")) and trim(rs("boardmaster"))<>"" then
			barr=split(rs("boardmaster"),"|")
			for i=0 to ubound(barr)
				if instr("|"&lcase(trim(request("oldboardmaster")))&"|","|"&barr(i)&"|")>0 then
					barr(i)=""
				end if
				if barr(i)<>"" then
					if boardmaster_1="" then 
						boardmaster_1=barr(i)
					else
						boardmaster_1=boardmaster_1&"|"&barr(i)
					end if
				end if
			next
		end if
		if boardmaster_1<>"" and cint(request("Board_Setting(50)"))=1 and trim(request("boardmaster"))<>"" then boardmaster_1=boardmaster_1&"|"
		if cint(request("Board_Setting(50)"))=1 then
			boardmaster_1=boardmaster_1&trim(request("boardmaster"))
		end if
		rs("boardmaster")=boardmaster_1
		iboard_Setting=request.Form("Board_Setting(0)") & "," & request.Form("Board_Setting(1)") & "," & request.Form("Board_Setting(2)") & "," & request.Form("Board_Setting(3)") & "," & request.Form("Board_Setting(4)") & "," & request.Form("Board_Setting(5)") & "," & request.Form("Board_Setting(6)") & "," & request.Form("Board_Setting(7)") & "," & request.Form("Board_Setting(8)") & "," & request.Form("Board_Setting(9)") & "," & request.Form("Board_Setting(10)") & "," & request.Form("Board_Setting(11)") & "," & request.Form("Board_Setting(12)") & "," & request.Form("Board_Setting(13)") & "," & request.Form("Board_Setting(14)") & "," & request.Form("Board_Setting(15)") & "," & request.Form("Board_Setting(16)") & "," & request.Form("Board_Setting(17)") & "," & request.Form("Board_Setting(18)") & "," & request.Form("Board_Setting(19)") & "," & request.Form("Board_Setting(20)") & "," & request.Form("Board_Setting(21)") & "," & request.Form("Board_Setting(22)") & "," & request.Form("Board_Setting(23)") & "," & request.Form("Board_Setting(24)") & "," & request.Form("Board_Setting(25)") & "," & request.Form("Board_Setting(26)") & "," & request.Form("Board_Setting(27)") & "," & request.Form("Board_Setting(28)") & "," & request.Form("Board_Setting(29)") & "," & request.Form("Board_Setting(30)") & "," & request.Form("Board_Setting(31)") & "," & request.Form("Board_Setting(32)") & "," & request.Form("Board_Setting(33)") & "," & request.Form("Board_Setting(34)") & "," & request.Form("Board_Setting(35)") & "," & request.Form("Board_Setting(36)") & "," & request.Form("Board_Setting(37)") & "," & request.Form("Board_Setting(38)") & "," & request.Form("Board_Setting(39)") & "," & request.Form("Board_Setting(40)") & "," & request.Form("Board_Setting(41)") & "," & request.Form("Board_Setting(42)") & "," & request.Form("Board_Setting(43)") & "," & request.Form("Board_Setting(44)") & "," & request.Form("Board_Setting(45)") & "," & request.Form("Board_Setting(46)") & "," & request.Form("Board_Setting(47)") & "," & request.Form("Board_Setting(48)") & "," & request.Form("Board_Setting(49)") & "," & request.Form("Board_Setting(50)") & "," & request.Form("Board_Setting(51)") & "," & request.Form("Board_Setting(52)") & "," & request.Form("Board_Setting(53)") & "," & request.Form("Board_Setting(54)") & "," & request.Form("Board_Setting(55)") & "," & request.Form("Board_Setting(56)") & "," & request.Form("Board_Setting(57)") & "," & request.Form("Board_Setting(58)") & "," & request.Form("Board_Setting(59)") & "," & request.Form("Board_Setting(60)") & "," & request.Form("Board_Setting(61)")
		rs("board_setting")=iboard_Setting
		rs.update
		rs.close
		set rs=nothing
		if (cint(request("Board_Setting(50)"))=1 and request("oldboardmaster")<>Request("boardmaster")) or (cint(request("Board_Setting(50)"))=0 and trim(request("oldboardmaster"))<>"") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)
		'============= 同学录 结束 ==================
		response.write "设置成功。<a href=admin_boardsetting.asp?editid="&request("editid")&">返回版面高级设置</a>"
	end if
end sub

sub addmaster(s,o,n)
	dim arr,pw,oarr
	dim classname,titlepic
	set rs=conn.execute("select title from usergroups where usergroupid=3")
	classname=rs(0)
	set rs=conn.execute("select titlepic from usertitle where usergroupid=3 order by Minarticle desc")
	if not (rs.eof and rs.bof) then
		titlepic=rs(0)
	end if
	randomize
	pw=Cint(rnd*9000)+1000
	arr=split(s,"|")
	oarr=split(o,"|")
	set rs=server.createobject("adodb.recordset")
	for i=0 to Ubound(arr)
		sql="select * from [user] where username='"& arr(i) &"'"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			rs.addnew
			rs("username")=arr(i)
			rs("userpassword")=md5(pw)
			rs("userclass")=classname
			rs("UserGroupID")=3
			rs("titlepic")=titlepic
			rs("userWealth")=Forum_user(0)
			rs("userep")=Forum_user(5)
			rs("usercp")=Forum_user(10)
			rs("userisbest")=0
			rs("userdel")=0
			rs("userpower")=0
			rs("lockuser")=0
			rs.update
			str=str&"你添加了以下用户：<b>" &arr(i) &"</b> 密码：<b>"& pw &"</b><br><br>"
		else
			if rs("UserGroupID")>3 then
				rs("userclass")=classname
				rs("UserGroupID")=3
				rs("titlepic")=titlepic
				rs.update
			end if
		end if
		rs.close
	next

	'判断原版主在其他版面是否还担任版主，如没有担任则撤换该用户职位
	if n=1 then
		dim iboardmaster
		dim UserGrade,article
		iboardmaster=false
		for i=0 to ubound(oarr)
			set rs=conn.execute("select boardmaster from board")
			do while not rs.eof
				if instr("|"&lcase(trim(rs("boardmaster")))&"|","|"&lcase(trim(oarr(i)))&"|")>0 then
					iboardmaster=true
					exit do
				end if
				rs.movenext
			loop
			if not iboardmaster then
				set rs=conn.execute("select userid,UserGroupID,article from [user] where username='"&trim(oarr(i))&"'")
				if not (rs.eof and rs.bof) then
					if rs(1)>2 then
						if not isnumeric(rs(2)) then article=0
						set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where Minarticle<="&rs(2)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
						if not (UserGrade.eof and UserGrade.bof) then
							conn.execute("update [user] set UserGroupID="&UserGrade(2)&",titlepic='"&UserGrade(1)&"',userclass='"&UserGrade(0)&"' where userid="&rs(0))
						end if
					end if
				end if
			end if
			iboardmaster=false
		next
	end if
	set rs=nothing
end sub
%>