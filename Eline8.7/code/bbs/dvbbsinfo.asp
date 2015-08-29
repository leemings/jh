<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%boardid=request("boardid")
if not isnumeric(boardid) or boardid="" or boardid=0 then
	boardid=0
end if
Stats="论坛变量信息"
call activeonline()
call nav()
call head_var(1,BoardDepth,0,0)%>
<BR>
<table cellpadding=6 cellspacing=1 align=center class=tableborder2>
	<tr><td class=tablebody1>论坛变量信息</td></tr>
	<tr><td class=tablebody1><a href=?view=Setting&boardid=<%=boardid%>>常规设置</a> | <a href=?view=info&boardid=<%=boardid%> >基本信息</a> | <a href=?view=Group&boardid=<%=boardid%>>用户组权限</a> | <a href=?view=css&boardid=<%=boardid%>>颜色CSS</a> | <a href=?view=pic&boardid=<%=boardid%>>论坛图片</a> | <a href=?view=boardpic&boardid=<%=boardid%>>论坛主体图标</a> | <a href=?view=TopicPic&boardid=<%=boardid%>>论坛帖子图标</a> | <a href=?view=userset&boardid=<%=boardid%>>用户设定</a> | <a href=?view=boardset&boardid=<%=boardid%>>分版设定</a></td></tr>
</table>
<%if  request("view")="Setting" then
	call Setting()
elseif request("view")="info" then
	call info()
elseif request("view")="Group" then
	call Group()
elseif request("view")="css" then
	call css()
elseif request("view")="pic" then
	call pic()
elseif request("view")="boardpic" then
	call boardpic()
elseif request("view")="TopicPic" then
	call TopicPic()
elseif request("view")="userset" then
	call userset()
elseif request("view")="boardset" then
	if boardid>0 then
		call boardset()
	else
		call Setting()
	end if
else
	call Setting()
end if
call footer()

sub Setting()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Setting()  常规设置信息</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Setting()  常规设置信息  当前值设置如下</td></tr>
    <%i=-1%>
    <tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(0) 服务器时差</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(1) 脚本超时时间(秒)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(2) 发送邮件组件  0－不支持 1－JMAIL 2－CDONTS  3－ASPEMAIL</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(3) 是否起用记录帖子阅读用户功能  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(4) 首页模式   0-新版样式 1-旧版样式</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(5) 首页显示论坛深度</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(6) 是否允许用户自定义头衔   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(7) 是否允许头像上传   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(8) 删除不活动用户时间(分钟)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(9) 版面语言</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(10) 新短消息弹出窗口  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(11) 帖子列表和浏览帖子除外的每页显示最多纪录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(12) 首页是否显示点歌列表(停用)   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(13) 首页是否显示回收站   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(14) 在线名单显示用户在线  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(15) 在线名单显示客人在线  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(16) 首页是否显示明星   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(17) 首页明星第一列的显示方式<BR>1-今日 2-本周 3-本月 4-本年 5-总帖数 6-最佳男星 7-最佳女星</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(18) 首页明星第二列的显示方式<BR>1-今日 2-本周 3-本月 4-本年 5-总帖数 6-最佳男星 7-最佳女星</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(19) 防刷新机制   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(20) 浏览刷新时间间隔(秒)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(21) 论坛当前状态</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(22) 同一IP注册间隔时间(秒)   0-不限</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(23) Email通知密码   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(24) 一个Email只能注册一个帐号 0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(25) 注册需要管理员认证  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(26) 允许同时在线数(人)   0-不限</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(27) 帖子中广告模式   0－静态广告 1－滚动新帖</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(28) 在线列表是否显示用户IP</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(29) 是否显示过生日会员  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(30) 是否显示页面执行时间  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(31) 快速登录位置   1－顶部 2－底部 0－不显示</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(32) 是否开启论坛门派  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(33) 在线资料列表显示当前位置  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(34) 在线资料列表显示登录和活动时间  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(35) 在线资料列表显示浏览器和操作系统  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(36) 在线资料列表显示来源  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(37) 是否允许新用户注册  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(38) 默认头像宽度(象素)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(39) 默认头像高度(象素)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(40) 最短用户名长度(字符)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(41) 最长用户名长度(字符)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(42) 允许个人签名   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(43) 是否显示周年纪念会员  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(44) 作为热门话题的最低人气值(回复数)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(45) 帖子中是否显示存款和股票  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(46) 开启短信欢迎新注册用户  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(47) 发送注册信息邮件  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(48) 编辑过的帖子显示"由xxx于yyy编辑"的信息</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(49) 管理员编辑后显示"由XXX编辑"的信息</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(50) 等待"由XXX编辑"信息显示的时间(分钟)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(51) 编辑帖子时限(分钟)   0-不限</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(52) 禁止的邮件地址</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(53) 允许用户使用头像   0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(54) 使用自定义头像的最少发帖数(帖)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(55) 允许从其他站点上传头像  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(56) 允许的最大头像文件大小(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(57) 最大头像宽度(象素)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(58) 最大头像高度(象素)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(59) 用户头衔最大长度(字符)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(60) 自定义头衔最少发帖数量限制(帖)   0-不限</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(61) 自定义头衔注册天数限制(天)   0-不限</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(62) 自定义头衔上面两个条件加在一起限制  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(63) 自定义头衔中要屏蔽的词语</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(64) 防刷新功能有效的页面</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(65) 用户签名是否开启UBB代码  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(66) 用户签名是否开启HTML代码 0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(67) 用户是否开启帖图标签  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(68) 用户排行列表个数(个)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(69) 是否定时开关论坛  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(70) 论坛开放时间</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(71) 导航菜单模式  0-顶部固定 1-左侧滑动</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(72) 帖子中会员属性显示方式  0-形象详细 1-形象简单 2-头像详细</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(73) 是否开启发帖限制  0-不限 1-限制主题数 2-限制回复数 3-限制帖子数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(74) 发帖限制比率</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(75) 首页是否显示“我最喜爱的论坛”  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(76) 版块图片是否淡入淡出  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(77) 首页Logo联盟论坛样式  0-不淡化不滚动 1-只淡化不滚动 2-只滚动不淡化 3-既滚动又淡化</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<%for i=78 to ubound(Forum_Setting)%>
			<tr>
				<td width="70%" class=tablebody2><li>Forum_Setting(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub 

sub info()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Info()  基本信息</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Info()   基本信息  当前值设置如下</td></tr>
    <%i=-1%>
    <tr>
			<td width="40%" class=tablebody2><li>Forum_Info(0) 论坛名称</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(1) 论坛的url</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(2) 主页名称</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(3) 主页URL</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(4) SMTP Server地址</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(5) 论坛管理员Email</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(6) 论坛首页Logo地址</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(7) 论坛图片目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(8) 论坛表情目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(9) 论坛所在时区</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(10) 发帖表情目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(11) 论坛头像目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(12) 论坛的建站日期</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<%for i=13 to ubound(Forum_Info)%>
			<tr>
				<td width="40%" class=tablebody2><li>Forum_Info(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Info(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub Group()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">GroupSetting()[reGroupSetting()]  用户组权限设置</th></tr>
    <tr><td class=tablebody1 colspan="2">GroupSetting()[reGroupSetting()]  用户组权限设置  当前值设置如下</td></tr>
    <%i=-1%>
    <tr>
			<td width="70%" class=tablebody2><li>GroupSetting(0)  可以浏览论坛   0－否 1－是</td>
			<td width="*" class=tablebody2>GroupSetting(0) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(0)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(1)  可以查看会员信息  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(2)  可以查看其他人发布的主题 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(3)  可以发布新主题   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(4)  可以回复自己的主题  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(5)  可以回复其他人的主题  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(6)  可以在论坛允许评分的时候参与评分0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(7)  可以上传附件   0－否 1－是 2-发帖可上传 3-回帖可上传</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(8)  可以发布新投票   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(9)  可以参与投票   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(10) 可以编辑自己的帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(11) 可以删除自己的帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(12) 可以移动自己的帖子到其他论坛 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(13) 以打开/关闭自己发布的主题 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(14) 可以搜索论坛   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(15) 可以使用“发送本页给好友”功能 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(16) 可以修改个人资料  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(17) 可以发布小字报   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(18) 可以删除其它人帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(19) 可以移动其它人帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(20) 可以打开/关闭其它人帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(21) 可以固顶/解除固顶帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(22) 可以奖励/惩罚发帖用户  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(23) 可以编辑其它人帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(24) 可以加入/解除精华帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(25) 可以发布公告   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(26) 可以管理公告   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(27) 可以管理小字报   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(28) 可以锁定/屏蔽/解除锁定用户 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(29) 可以删除用户1－10天内所发帖子 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(30) 可以查看来访IP及来源  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(31) 可以限定IP来访   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(32) 可以发送短信   0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(33) 最多发送用户</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(34) 短信内容大小限制(字节)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(35) 信箱大小限制(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(36) 是否有审核帖子的权限 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(37) 是否有进入隐含论坛的权限 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(38) 可以进行帖子总固顶操作 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(39) 可以浏览论坛事件  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(40) 一次最多上传文件个数(个)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(41) 可以浏览精华帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(42) 可以管理用户权限  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(43) 可以奖励/惩罚用户  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(44) 上传文件大小限制(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(45) 可以批量删除帖子（前台） 0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(46) 发布小字报所需金钱(元)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(47) 参与评分所需金钱(元)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(48) 有论坛文件管理权限  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(49) 可以浏览论坛展区  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(50) 一天最多上传文件个数(个)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(51) 有帖子主题颜色权限  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<%for i=52 to ubound(GroupSetting)%>
			<tr>
				<td width="70%" class=tablebody2><li>GroupSetting(<%=i%>) </td>
				<td width="*" class=tablebody2>GroupSetting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub css()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Body()  关于颜色CSS的变量</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Body()  关于颜色CSS的变量  当前值设置如下</td></tr>
		<%i=-1%>
    <tr>
			<td width="60%" class=tablebody2><li>Forum_Body(0) 表格边框（体）CSS定义一 调用：Class=TableBorder1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(1) 表格边框（体）CSS定义二 调用：Class=TableBorder2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(2) 标题表格CSS定义一（深背景） 调用：th</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(3) 标题表格CSS定义二（浅背景） 调用：Class=TableTitle2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(4) 表格体CSS定义一 调用：Class=TableBody1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(5) 表格体颜色二(1和2颜色在显示中穿插) 调用：Class=TableBody2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(6) 论坛连接总的CSS定义</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(7) 标题表格字体连接CSS定义（深背景） 调用：id=TableTitleLink</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(8) 警告提醒语句的颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(9) 显示帖子的时候，相关帖子，转发帖子，回复等的颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(10) 论坛表格总的CSS定义</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(11) 论坛BODY标签</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(12) 表格宽度</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(13)  </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(14) 首页连接颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(15) 一般用户名称字体颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(16) 一般用户名称上的光晕颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(17) 版主名称字体颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(18) 版主名称上的光晕颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(19) 管理员名称字体颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(20) 管理员名称上的光晕颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(21) 贵宾名称字体颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(22) 贵宾名称上的光晕颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(23) 顶部菜单表格CSS定义(Logo & Banner下方) 调用：Class=TopLighNav</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(24) 顶部菜单表格CSS定义(Logo & Banner上方) 调用：Class=TopDarkNav</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(25) Body的CSS定义</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(26) 顶部菜单表格CSS定义(导航菜单) 调用：Class=TopLighNav1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(27) 表格边框颜色定义 在下面定义一和二CSS定义控制不到的地方，最好保持和边框CSS定义一中同样的颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(28) 顶部菜单表格CSS定义(Logo & banner部分) 调用：Class=TopLighNav2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<%for i=29 to ubound(Forum_Body)%>
			<tr>
				<td width="60%" class=tablebody2><li>Forum_Body(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Body(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub pic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr><th height="24" colspan="2">Forum_Pic()  论坛图片变量</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Pic()  论坛图片变量  当前值设置如下</td></tr>
    <%i=-1%>
    <tr>
			<td width="30%" class=tablebody2><li>Forum_Pic(0) 论坛管理员</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(1) 论坛版主</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(2) 论坛贵宾</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(3) 普通会员</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(4) 客人或隐身会员</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(5) 突出显示自己的颜色</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(6) 常规论坛－－无新帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(7) 常规论坛－－有新帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(8) 版面列表版面说明前面的图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(9) 版面列表中今日发帖数图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(10) 版面列表中总发帖数图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(11) 版面列表中总主题数图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(12) 论坛导航图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(13) 帖子列表上方版主列表图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(14) 认证论坛－－无新帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(15) 认证论坛－－有新帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(16) 超级版主</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<%for i=17 to ubound(Forum_Pic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_Pic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Pic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
			</tr>
		<%next%>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>
<%end sub 

sub boardpic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr><th height="24" colspan="2">Forum_BoardPic()  论坛主体图标</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_BoardPic()  论坛主体图标  当前值设置如下</td></tr>
    <%i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(0) 联盟论坛</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(1) 发表帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(2) 发表投票</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(3) 小字报</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(4) 回复帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(5) 帮助</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(6) 返回首页</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(7) 返回上级目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(8) 当前目录</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(9) 新的短消息</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(10) 我发表的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(11) 树形浏览帖子模式</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(12) 平板形浏览帖子模式</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(13) 下一篇帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(14) 上一篇帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(15) 刷新浏览</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(16) 论坛公告</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<%for i=17 to ubound(Forum_BoardPic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_BoardPic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
			</tr>
		<%next
		i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(0) 打开的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(1) 热门的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(2) 锁定的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(3) 固顶的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(4) 精华的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(5) 版面列表中最后发帖</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(6) 帖子列表中分页图标</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(7) 刷新</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(8) 短信提示音文件名</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(9) 总固顶的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(10) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(11) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(12) 投票的主题</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<%for i=13 to ubound(Forum_StatePic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_StatePic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
			</tr>
		<%next
		i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(0) 粗体</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(1) 斜体</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(2) 下划线</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(3) 居中</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(4) URL连接</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(5) Email地址</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(6) 贴图</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(7) 贴Flash</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(8) 贴ShockWave</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(9) 贴RM影片</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(10) 贴Media Player影片</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(11) 贴QuickTime影片</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(12) 引用文字</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(13) 飞行字</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(14) 移动字</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(15) 发光字</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(16) 阴影字</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(17) 心情列表</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(18) 背景音乐</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<%for i=19 to ubound(Forum_UBB)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_UBB(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_UBB(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub

sub topicpic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_TopicPic()  论坛帖子图标</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_TopicPic()  论坛帖子图标  当前值设置如下</td></tr>
    <%i=-1%>
    <tr>
			<td width="30%" class=tablebody2><li>Forum_TopicPic(0) 保存当页为文件</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(1) 报告本帖给版主</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(2) 显示可打印的版本</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(3) 把本帖打包邮递</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(4) 把本帖加入论坛收藏</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(5) 发送本页面给朋友</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(6) 把本帖加入IE收藏</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(7) 发送短信给好友</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(8) 搜索用户发表帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(9) 浏览用户信息</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(10) 用户邮箱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(11) 用户OICQ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(12) 用户ICQ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(13) 用户MSN</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(14) 用户主页</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(15) 引用帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(16) 编辑帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(17) 删除帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(18) 复制帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(19) 加入精华帖子</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(20) 用户IP</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(21) 加为好友</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(22) 快速回复</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<%for i=23 to ubound(Forum_TopicPic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_TopicPic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub

sub userset()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_User() 用户设定</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_User() 用户设定  当前值设置如下</td></tr>
    <%i=-1%>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(0) 注册金钱数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(1) 发帖增加金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(2) 跟帖增加金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(3) 删帖减少金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(4) 登录增加金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(5) 注册经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(6) 发帖增加经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(7) 跟帖增加经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(8) 删帖减少经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(9) 登录增加经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(10) 注册魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(11) 发帖增加魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(12) 跟帖增加魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(13) 删帖减少魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(14) 登录增加魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(15) 精华增加金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(16) 精华增加魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(17) 精华增加经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(18) 普通会员给所有会员点歌的价格(停用)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(19) 普通会员给单一会员点歌的价格(停用)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(20) VIP会员给所有会员点歌的价格(停用)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(21) VIP会员给单一会员点歌的价格(停用)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(22) 普通会员修改一次头像的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(23) VIP会员修改一次头像的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(24) 普通会员修改一次签名的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(25) VIP会员修改一次签名的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(26) 普通会员发送短信的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(27) VIP会员发送短信的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(28) 普通会员回复后短信通知的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(29) VIP会员回复后短信通知的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(30) 普通会员回复后邮件通知的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(31) VIP会员回复后邮件通知的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(32) 普通会员短信引起注意的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(33) VIP会员短信引起注意的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(34) 普通会员上传头像的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(35) VIP会员上传头像的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(36) 普通会员上传文件的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(37) VIP会员上传文件的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(38) 普通会员上传照片的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(39) VIP会员上传照片的价格</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(40) 投票增加金钱</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(41) 投票增加经验值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(42) 投票增加魅力值</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<%for i=43 to ubound(Forum_User)%>
			<tr>
				<td width="60%" class=tablebody2><li>Forum_User(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_User(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub

sub boardset()
	dim i
	%><table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Board_Setting() 分版论坛设定</th></tr>
		<tr><td class=tablebody1 colspan="2">Board_Setting() 分版论坛设定  当前值设置如下</td></tr>
		<%i=-1%>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(0) 论坛开放/锁定</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(1) 论坛隐含/不隐含（用户组中设定是否可进入隐含论坛）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(2) 论坛加密/不加密（加密论坛需要设定可进入用户）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(3) 论坛审核不开放/开放</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(4) 发帖模式简单/高级</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(5) HTML支持开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(6) UBB功能开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(7) 帖图标签开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(8) 表情标签开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(9) 多媒体标签开启/关闭（包括Flash,RM,AVI等）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(10) 金钱帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(11) 积分帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(12) 魅力帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(13) 威望帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(14) 文章帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(15) 回复帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(16) 帖子内容最大字节数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(17) 发帖后返回（首页/论坛/帖子）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(18) 允许同时在线数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(19) 上传文件类型</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(20) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(21) 启动版面定时开放功能</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(22) 论坛开放时间（0－24点）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(23) 版面语言（简体中文/繁体中文/英文）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(24) 论坛风格（讨论区风格/默认风格）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(25) 讨论区风格预览字数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(26) 帖子列表每页帖子数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(27) 浏览帖子每页帖子数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(28) 帖子正文字号</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(29) 帖子正文行间距</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(30) 防灌水机制</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(31) 每次发帖时间间隔(秒)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(32) 最多投票项目</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(33) 主版主可以增删副版主</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(34) 主版主可以修改CSS设置</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(35) 所有版主可以修改CSS设置</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(36) 是否对论坛事件中的操作者保密（对管理员无效）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(37) 论坛默认读取帖子量（按所需时间内）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(38) 默认排序（最后回复时间/标题/发帖人/回复数/浏览数/发表时间）</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(39) 是否采用简洁风格列表  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(40) 是否继承上级版主权限  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(41) 简洁风格列表一行版面数</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(42) 是否显示下级论坛帖子  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(43) 作为分类论坛是否不允许发帖  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(44) 是否开放发帖机遇功能  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(45) 发帖机遇的默认状态  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(46)  是否开放发帖后短信回复功能  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(47) 发帖后短信回复的默认状态  0－关闭 1－打开</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(48) 是否开放短信引起注意功能  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(49) 是否开放版主评分功能  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(50) 是否同学录论坛  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(51) 是否VIP论坛  0－否 1－是</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(52) 是否性别论坛  0－普通 1－女生 2-男生</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(53) 秘密帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(54) 会员帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(55) 定人帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(56) 性别帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(57) 高级帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(58) 年龄帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(59) 门派帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(60) 显示帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(61) 隐藏帖开启/关闭</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<%for i=62 to ubound(Board_Setting)%>
			<tr>
				<td width="70%" class=tablebody2><li>Board_Setting(<%=i%>) </td>
				<td width="*" class=tablebody2>Board_Setting(<%=i%>) 当前值为：<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub%>