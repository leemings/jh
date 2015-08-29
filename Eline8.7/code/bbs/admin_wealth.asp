<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="23"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
		call savegrade()
		else
		call grade()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub grade()
dim sel
%>
<form method="POST" action=admin_wealth.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛风格模板中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛设置中，可多选<BR>3、如果您想在一个版面引用别的版面或模板的配置，只要点击该版面或模板名称，保存的时候选择要保存到的版面名称或模板名称即可。</font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2><B>当前使用主模板</B>（可将设置保存到下列模板中）</th>
</tr>
<tr> 
<td width="100%" colspan=2 class="forumrow" height=25><input type="checkbox" name="checkall" checked value="1">&nbsp;全选：如果要使下面设置对所有模板生效，请全选，建议全选
</td>
<tr>
<td colspan=2 class=forumrow>
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
	Forum_user=split(rs("Forum_user"),",")
	uForum_user=ubound(Forum_user)
	if uForum_user<39 then
		redim preserve Forum_user(39)
		Forum_user(18)=100
		Forum_user(19)=50
		Forum_user(20)=50
		Forum_user(21)=25
		Forum_user(22)=20
		Forum_user(23)=10
		Forum_user(24)=40
		Forum_user(25)=20
		Forum_user(26)=4
		Forum_user(27)=2
		Forum_user(28)=40
		Forum_user(29)=20
		Forum_user(30)=0
		Forum_user(31)=0
		Forum_user(32)=10
		Forum_user(33)=5
		Forum_user(34)=2
		Forum_user(35)=1
		Forum_user(36)=2
		Forum_user(37)=1
		Forum_user(38)=2
		Forum_user(39)=1
	end if
	if uForum_user<42 then
		redim preserve Forum_user(42)
		Forum_user(40)=5
		Forum_user(41)=1
		Forum_user(42)=1
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
	Forum_user=split(rs("Forum_user"),",")
	uForum_user=ubound(Forum_user)
	if uForum_user<39 then
		redim preserve Forum_user(39)
		Forum_user(18)=100
		Forum_user(19)=50
		Forum_user(20)=50
		Forum_user(21)=25
		Forum_user(22)=20
		Forum_user(23)=10
		Forum_user(24)=40
		Forum_user(25)=20
		Forum_user(26)=4
		Forum_user(27)=2
		Forum_user(28)=40
		Forum_user(29)=20
		Forum_user(30)=0
		Forum_user(31)=0
		Forum_user(32)=10
		Forum_user(33)=5
		Forum_user(34)=2
		Forum_user(35)=1
		Forum_user(36)=2
		Forum_user(37)=1
		Forum_user(38)=2
		Forum_user(39)=1
	end if
	if uForum_user<42 then
		redim preserve Forum_user(42)
		Forum_user(40)=5
		Forum_user(41)=1
		Forum_user(42)=1
	end if
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_wealth.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=Forumrow><B>当前使用该模板的分论坛</B><BR>以下分论坛将使用本模板设置，如需修改论坛使用模板，可到论坛管理中重新设置<BR>
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
<th height="23" colspan="2">用户金钱设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册金钱数</td>
<td width="60%" class=Forumrow><input type="text" name="wealthReg" size="35" value="<%=Forum_user(0)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登录增加金钱</td>
<td width="60%" class=Forumrow><input type="text" name="wealthLogin" size="35" value="<%=Forum_user(4)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加金钱</td>
<td width="60%" class=Forumrow><input type="text" name="wealthAnnounce" size="35" value="<%=Forum_user(1)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加金钱</td>
<td width="60%" class=Forumrow><input type="text" name="wealthReannounce" size="35" value="<%=Forum_user(2)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加金钱</td>
<td width="60%" class=Forumrow><input type="text" name="BestWealth" size="35" value="<%=Forum_user(15)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少金钱</td>
<td width="60%" class=Forumrow><input type="text" name="wealthDel" size="35" value="<%=Forum_user(3)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>投票增加金钱</td>
<td width="60%" class=Forumrow><input type="text" name="wealthVote" size="35" value="<%=Forum_user(40)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >用户经验设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epReg" size="35" value="<%=Forum_user(5)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登录增加经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epLogin" size="35" value="<%=Forum_user(9)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epAnnounce" size="35" value="<%=Forum_user(6)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epReannounce" size="35" value="<%=Forum_user(7)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加经验值</td>
<td width="60%" class=Forumrow><input type="text" name="bestuserep" size="35" value="<%=Forum_user(17)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epDel" size="35" value="<%=Forum_user(8)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>投票增加经验值</td>
<td width="60%" class=Forumrow><input type="text" name="epVote" size="35" value="<%=Forum_user(41)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >用户魅力设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>注册魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpReg" size="35" value="<%=Forum_user(10)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>登录增加魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpLogin" size="35" value="<%=Forum_user(14)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>发帖增加魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpAnnounce" size="35" value="<%=Forum_user(11)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>跟帖增加魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpReannounce" size="35" value="<%=Forum_user(12)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>精华增加魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="bestusercp" size="35" value="<%=Forum_user(16)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>删帖减少魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpDel" size="35" value="<%=Forum_user(13)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>投票增加魅力值</td>
<td width="60%" class=Forumrow><input type="text" name="cpVote" size="35" value="<%=Forum_user(42)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >服务收费价格设定</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员给全体会员点歌</td>
<td width="60%" class=Forumrow><input type="text" name="dgallDel" size="35" value="<%=Forum_user(18)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员给单一会员点歌</td>
<td width="60%" class=Forumrow><input type="text" name="dgDel" size="35" value="<%=Forum_user(19)%>">&nbsp;元/人</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员给全体会员点歌</td>
<td width="60%" class=Forumrow><input type="text" name="dgvipallDel" size="35" value="<%=Forum_user(20)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员给单一会员点歌</td>
<td width="60%" class=Forumrow><input type="text" name="dgvipDel" size="35" value="<%=Forum_user(21)%>">&nbsp;元/人</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员修改一次头像</td>
<td width="60%" class=Forumrow><input type="text" name="faceDel" size="35" value="<%=Forum_user(22)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员修改一次头像</td>
<td width="60%" class=Forumrow><input type="text" name="facevipDel" size="35" value="<%=Forum_user(23)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员修改一次签名</td>
<td width="60%" class=Forumrow><input type="text" name="signDel" size="35" value="<%=Forum_user(24)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员修改一次签名</td>
<td width="60%" class=Forumrow><input type="text" name="signvipDel" size="35" value="<%=Forum_user(25)%>">&nbsp;元/次</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员发送短信</td>
<td width="60%" class=Forumrow><input type="text" name="mesDel" size="35" value="<%=Forum_user(26)%>">&nbsp;元/人</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员发送短信</td>
<td width="60%" class=Forumrow><input type="text" name="mesvipDel" size="35" value="<%=Forum_user(27)%>">&nbsp;元/人</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员回复后短信通知</td>
<td width="60%" class=Forumrow><input type="text" name="msgDel" size="35" value="<%=Forum_user(28)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员回复后短信通知</td>
<td width="60%" class=Forumrow><input type="text" name="msgvipDel" size="35" value="<%=Forum_user(29)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员回复后邮件通知</td>
<td width="60%" class=Forumrow><input type="text" name="emailDel" size="35" value="<%=Forum_user(30)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员回复后邮件通知</td>
<td width="60%" class=Forumrow><input type="text" name="emailvipDel" size="35" value="<%=Forum_user(31)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员短信引起注意</td>
<td width="60%" class=Forumrow><input type="text" name="notifyDel" size="35" value="<%=Forum_user(32)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员短信引起注意</td>
<td width="60%" class=Forumrow><input type="text" name="notifyvipDel" size="35" value="<%=Forum_user(33)%>">&nbsp;元/帖</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员上传头像</td>
<td width="60%" class=Forumrow><input type="text" name="upfaceDel" size="35" value="<%=Forum_user(34)%>">&nbsp;元/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员上传头像</td>
<td width="60%" class=Forumrow><input type="text" name="upfacevipDel" size="35" value="<%=Forum_user(35)%>">&nbsp;元/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员上传文件</td>
<td width="60%" class=Forumrow><input type="text" name="upfileDel" size="35" value="<%=Forum_user(36)%>">&nbsp;元/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员上传文件</td>
<td width="60%" class=Forumrow><input type="text" name="upfilevipDel" size="35" value="<%=Forum_user(37)%>">&nbsp;元/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>普通会员上传照片</td>
<td width="60%" class=Forumrow><input type="text" name="upphotoDel" size="35" value="<%=Forum_user(38)%>">&nbsp;元/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP会员上传照片</td>
<td width="60%" class=Forumrow><input type="text" name="upphotovipDel" size="35" value="<%=Forum_user(39)%>">&nbsp;元/KB</td>
</tr>
<tr height=23> 
<td colspan=2 class=Forumrow align=center><input type="submit" name="Submit" value="提 交"></td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>请选择保存的模板名称"
else
Forum_user=request.form("wealthReg") & "," & request.form("wealthAnnounce") & "," & request.form("wealthReannounce") & "," & request.form("wealthDel") & "," & request.form("wealthLogin") & "," & request.form("epReg") & "," & request.form("epAnnounce") & "," & request.form("epReannounce") & "," & request.form("epDel") & "," & request.form("epLogin") & "," & request.form("cpReg") & "," & request.form("cpAnnounce") & "," & request.form("cpReannounce") & "," & request.form("cpDel") & "," & request.form("cpLogin") & "," & request.form("BestWealth") & "," & request.form("BestuserCP") & "," & request.form("BestuserEP") & "," & request.form("dgalldel") & "," & request.form("dgdel") & "," & request.form("dgvipalldel") & "," & request.form("dgvipdel") & "," & request.form("facedel") & "," & request.form("facevipdel") & "," & request.form("signdel") & "," & request.form("signvipdel") & "," & request.form("mesdel") & "," & request.form("mesvipdel") & "," & request.form("msgdel") & "," & request.form("msgvipdel") & "," & request.form("emaildel") & "," & request.form("emailvipdel") & "," & request.form("notifydel") & "," & request.form("notifyvipdel") & "," & request.form("upfacedel") & "," & request.form("upfacevipdel") & "," & request.form("upfiledel") & "," & request.form("upfilevipdel") & "," & request.form("upphotodel") & "," & request.form("upphotovipdel") & "," & request.form("wealthVote") & "," & request.form("epVote") & "," & request.form("cpVote")
if request("checkall")="1" then
sql="update config set Forum_user='"&Forum_user&"'"
conn.execute(sql)
else
if request("skinid")<>"" then
sql = "update config set Forum_user='"&Forum_user&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
end if
response.write "设置成功。"
end if
end sub
%>