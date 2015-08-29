<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="z_fastpost_conn.asp"-->
<%dim AnnounceID
dim replyID
dim username
dim rs_old
dim old_user
dim con,content
dim topic
dim olduser,oldpass
dim totalusetable
dim CanEditPost
dim signflag
dim mailflag
dim msgflag
Dim ChenTopicColor
dim Speak

CanEditPost=False
if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
		end if
	end if
end if
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		dim sexshow
		if cint(Board_Setting(52))=1 then
			sexshow="女生"
		elseif cint(Board_Setting(52))=2 then
			sexshow="男生"
		end if
		if chksexlogin(cint(Board_Setting(52)),membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>"&sexshow&"论坛版面</font>，请确认您的性别是否符合。"
			founderr=true
		end if
	end if
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	AnnounceID=request("id")
end if
if request("replyID")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
elseif not isInteger(request("replyID")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
else
	replyID=request("replyID")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
end if
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发帖。"
	founderr=true
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
end if
stats=boardtype & "编辑帖子"
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
	dim CanTopicColor,isTopic
	set rs=conn.execute("select PostTable,ChenTopicColor from topic where BoardID="&boardid&" and TopicID="&AnnounceID)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到相应的帖子。"
		Founderr=true
		exit sub
	else
		TotalUseTable=rs(0)
		ChenTopicColor=rs(1)
	end if
	sql="select b.username,b.topic,b.body,b.dateandtime,u.usergroupID,b.signflag,b.emailflag,b.speak,b.parentid from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.rootid="&AnnounceID&" and b.BoardID="&boardid&" and  b.AnnounceID="&replyID
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到相应的帖子。"
		Founderr=true
		exit sub
	else
		if rs("parentid")=0 then
			isTopic=True
		else
			isTopic=False
		end if
		topic=rs("topic")
		con=rs("body")
		signflag=rs("signflag")
		mailflag=rs("emailflag")
		if mailflag>1 then
			msgflag=1
		else
			msgflag=0
		end if
		mailflag=(mailflag mod 2)
		speak=rs("speak")
		old_user=rs("username")
		if Clng(forum_setting(51))>0 and not (master or boardmaster or superboardmaster) then
			if Datediff("s",rs("dateandtime"),Now())>Clng(forum_setting(51))*60 then
			Errmsg=Errmsg+"<br>"+"<li>系统编辑帖子时限为"&forum_setting(51)&"分钟，而从您发表帖子到现在已经有"&Datediff("s",rs("dateandtime"),Now())\60&"分钟了。"
				founderr=true
				exit sub
			end if
		end if
		if rs("username")=membername then
			if Cint(GroupSetting(10))=0 then
				Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛编辑自己帖子的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
				founderr=true
				exit sub
				CanEditPost=False
			else
				CanEditPost=True
			end if
		else
			if (master or superboardmaster or boardmaster) and Cint(GroupSetting(23))=1 then
				CanEditPost=True
			else
				CanEditPost=False
			end if
			if UserGroupID>3 and Cint(GroupSetting(23))=1 then
				CanEditPost=true
			end if
			if Cint(GroupSetting(23))=1 and FoundUserPer then
				CanEditPost=True
			elseif Cint(GroupSetting(23))=0 and FoundUserPer then
				CanEditPost=False
			end if
			if UserGroupID<3 and UserGroupID=rs("UserGroupID") then
				Errmsg=Errmsg+"<br>"+"<li>同等级用户不能修改。"
				Founderr=true
				exit sub
			elseif UserGroupID=3 and UserGroupID=rs("UserGroupID") then
				if boardmaster then
					dim rss,userboardmaster
					sql="select BoardMaster from board where boardid="&boardid
					set rss=server.createobject("adodb.recordset")
					rss.open sql,conn,1,1
					userboardmaster=rss(0)
					rss.close
					set rss=nothing
					if instr(1,"|"&lcase(trim(userboardmaster))&"|","|"&lcase(trim(rs("username")))&"|")>0 then
						Errmsg=Errmsg+"<br>"+"<li>同等级用户不能修改。"
						Founderr=true
						exit sub
					end if
				end if
			elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
				Errmsg=Errmsg+"<br>"+"<li>不能修改等级比您高的用户的帖子。"
				Founderr=true
				exit sub
			end if
			if not CanEditPost then
				Errmsg=Errmsg+"<br>"+"<li>您没有足够的权限编辑本帖子，请和管理员联系。"
				Founderr=true
				exit sub
			end if
		end if
	end if
	set rs=nothing
	if con<>"" then
		content=replace(con,"<br>",chr(13))
		content=replace(content,"&nbsp;","")
		'content=content+chr(13)
	end if
	'判断用户是否有帖子主题颜色权限
	if isTopic then
		if (master or superboardmaster or boardmaster) and Cint(GroupSetting(51))=1 then
			CanTopicColor=true
		else
			CanTopicColor=false
		end if
		if Cint(GroupSetting(51))=1 and UserGroupID>3 then
			CanTopicColor=true
		end if
		if FoundUserPer and Cint(GroupSetting(51))=1 then
			CanTopicColor=true
		elseif FoundUserPer and Cint(GroupSetting(51))=0 then
			CanTopicColor=false
		end if
	end if%>
	<script src="inc/ubbcode.js"></script>
	<form action="SaveditAnnounce.asp?boardID=<%=boardID%>&replyID=<%=replyID%>&ID=<%=announceID%>" method=POST name=frmAnnounce>
  <input type=hidden name=followup value=<%=AnnounceID%>>
  <input type=hidden name="star" value="<%=request("star")%>">
  <input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
		<tr>
			<th width=100% colspan=2>编辑帖子</th>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>用户名</b></td>
			<td width=80% class=tablebody2><input name=username value=<%=htmlencode(old_user)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>您没有注册？</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody1><b>密码</b></td>
			<td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=lostpass.asp>忘记密码？</a>版主编辑不需要密码</td>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>主题标题</b>
				<SELECT name=font onchange=DoTitle2(this.options[this.selectedIndex].value)>
				<OPTION selected value="">选择话题</OPTION> <OPTION value=[原创]>[原创]</OPTION> 
				<OPTION value=[转帖]>[转帖]</OPTION> <OPTION value=[灌水]>[灌水]</OPTION> 
				<OPTION value=[讨论]>[讨论]</OPTION> <OPTION value=[求助]>[求助]</OPTION> 
				<OPTION value=[推荐]>[推荐]</OPTION> <OPTION value=[公告]>[公告]</OPTION> 
				<OPTION value=[注意]>[注意]</OPTION> <OPTION value=[帖图]>[帖图]</OPTION>
				<OPTION value=[建议]>[建议]</OPTION> <OPTION value=[下载]>[下载]</OPTION>
				<OPTION value=[分享]>[分享]</OPTION>
				</SELECT>
			</td>
			<td width=80% class=tablebody2><input name=subject size=70 maxlength=80 value=<%=htmlencode(Topic)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>不得超过 25 个汉字或50个英文字符</td>
		</tr>
		<tr height=20> 
			<td width=20% class=tablebody1><b>标题颜色</b></td>
			<td width=80% class=tablebody1><input type=hidden name=ChenTopicColor value="<%=ChenTopicColor%>"> <%
				if not CanTopicColor then
					%><font color="<%=forum_body(8)%>">您没有权限修改帖子主题颜色或该帖子不是主题</font> <%
				else
					%><SCRIPT language=JavaScript src="inc/z_color.js"></SCRIPT>
					<SCRIPT language=JavaScript>
					var height0 = 10;		// define the height of the color bar
					var pas0 = 36;			// define the number of color in the color bar
					var width0=2;				//define the width of the color bar here (for forum with little width put 1)
					</SCRIPT>
					<TABLE id=HeadColorPanel cellSpacing=0 cellPadding=0 align=left border=0 >
					<TBODY>
						<TR><TD id=ColorHead onclick="this.bgColor='';MM_findObj('ChenTopicColor').value=''" vAlign=center align=middle <%if not isnull(ChenTopicColor) and ChenTopicColor<>"" then%>bgcolor="<%=ChenTopicColor%>"<%end if%>><SCRIPT language=JavaScript>document.write('<IMG height='+height0+' width=20 border=1></TD>');</SCRIPT></TD><SCRIPT language=JavaScript>rgb1(pas0,width0,height0)</SCRIPT></TR>
					</TBODY>
					</TABLE><%
				end if
			%></td>
		</tr>
		<tr>
			<td width=20% valign=top class=tablebody2><b>当前心情</b><br><li>将放在帖子的前面</td>
			<td width=80% class=tablebody2><%
				for i=0 to Forum_PostFaceNum-1
					%><input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;<%
					if i>0 and ((i+1) mod 9=0) then response.write "<br>"
				next
			%></td>
		</tr>
		<tr>
			<td width=20% valign=top class=tablebody1><b>内容</b><br><!--#include file="z_ContentFlag.asp"--></td>
			<td width=80% class=tablebody1><%if Cint(Board_Setting(4))=1 then%><!--#include file="inc/getubb.asp"--><%end if%><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL  onkeydown=ctlent()><%=server.htmlencode(content)%></textarea><!--#include file="z_speak.asp"--></td>
		</tr>
		<tr>
			<td class=tablebody2 valign=top colspan=2 style="table-layout:fixed; word-break:break-all"><b>点击表情图即可在帖子中加入相应的表情</B><br>&nbsp;<%
				dim ii
				for i=0 to 14
					if len(i)=1 then ii="0" & i else ii=i
					response.write "<img src="""&Forum_info(10)&Forum_emot(i)&""" border=0 onclick=""insertsmilie('[em"&ii&"]')"" style=""CURSOR: hand"">&nbsp;"
				next
			%></td>
		</tr>
		<tr>
			<td valign=top class=tablebody1><b>选项</b></td>
			<td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes <%if signflag then%>checked<%end if%>>是否显示您的签名？<%
				if board_setting(46)=1 then
					%><br><input type=checkbox name=msgflag value=yes <%if msgflag=1 then%>checked<%end if%>>有回复时使用短信通知您？(<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(28)
					else
						response.write forum_user(29)
					end if
					%>元，您现在有现金<%=mymoney%>元</font>)<%
				else
					%><input type=hidden name=msgflag value=no><%
				end if
				%><br><input type=checkbox name=emailflag value=yes <%if mailflag=1 then%>checked<%end if%>>有回复时使用邮件通知您？(<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(30)
				else
					response.write forum_user(31)
				end if
			%>元，您现在有现金<%=mymoney%>元</font>)</td>
		</tr>
		<tr>
			<td class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="发 表" name=Submit onclick="if(checkbbs(<%if isTopic then%>true<%else%>false<%end if%>)){return true;} return false;"> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value="清 除"></td>
		</tr>
	</table>
	</form>
	<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
		<input type=hidden name=title value=><input type=hidden name=body value=>
	</form>
<%end sub%>