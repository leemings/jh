<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="z_fastpost_conn.asp"-->
<%if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats=boardtype&"发表投票"
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
	dim CanTopicColor
	if (Cint(Board_Setting(43))=1 or Cint(Board_Setting(0))=1) and not (master or superboardmaster or boardmaster) then
		Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发帖。"
		founderr=true
		exit sub
	end if
	if cint(Board_Setting(1))=1 then
		if Cint(GroupSetting(37))=0 then
			Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
			founderr=true
			exit sub
		end if
	end if
	if cint(Board_Setting(51))=1 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkviplogin(membername)=false then
				Errmsg=Errmsg+"<br>"+"<li>本论坛版面为<font color=red>VIP会员专用版面</font>，请确认您的属性是否符合。"
				founderr=true
				exit sub
			end if
		end if
	end if
	if cint(Board_Setting(52))<>0 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为特殊论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
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
				exit sub
			end if
		end if
	end if
	if Cint(GroupSetting(8))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛发表投票的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登录</a>并确认您的用户名已经得到管理员的认证后进入。"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
				Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
				founderr=true
				exit sub
			end if
		end if
	end if
	'判断用户是否有帖子主题颜色权限
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
	end if%>
	<script src="inc/ubbcode.js"></script>
	<form action=Savevote.asp?boardID=<%=request("boardid")%> method=POST onSubmit=submitonce(this) name=frmAnnounce>
	<input type="hidden" name="upfilerename">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
		<tr>
			<th width=100% colspan=2 align=left>&nbsp;&nbsp;发表新投票<%if isaudit=1 then%>（本版面所有帖子都要经过管理员审核方可发表）<%end if%>  </th>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>用户名</b></td>
			<td width=80% class=tablebody2><input name=username value=<%=membername%> class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>您没有注册？</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody1><b>密码</b></td>
			<td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> class=FormClass><font color=<%=Forum_body(8)%>>&nbsp;&nbsp;<b>*</b></font><a href=lostpass.asp>忘记密码？</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>主题标题</b>&nbsp;<SELECT name=font onchange=DoTitle2(this.options[this.selectedIndex].value)>
				<OPTION selected value="">选择话题</OPTION> <OPTION value=[原创]>[原创]</OPTION> 
				<OPTION value=[转帖]>[转帖]</OPTION> <OPTION value=[灌水]>[灌水]</OPTION> 
				<OPTION value=[讨论]>[讨论]</OPTION> <OPTION value=[求助]>[求助]</OPTION> 
				<OPTION value=[推荐]>[推荐]</OPTION> <OPTION value=[公告]>[公告]</OPTION> 
				<OPTION value=[注意]>[注意]</OPTION> <OPTION value=[帖图]>[帖图]</OPTION>
				<OPTION value=[建议]>[建议]</OPTION> <OPTION value=[下载]>[下载]</OPTION>
				<OPTION value=[分享]>[分享]</OPTION></SELECT>
			</td>
			<td width=80% class=tablebody2><input name=subject size=50 maxlength=80 class=FormClass>&nbsp;<select name=votetimeout size=1>
				<option value=0>过期时间</option>
				<option value=0>永不过期</option>
				<option value=1>一天</option>
				<option value=3>三天</option>
				<option value=7>一周</option>
				<option value=15>半月</option>
				<option value=30>一月</option>
				<option value=90>三月</option>
				<option value=180>半年</option>
				</select>&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>不得超过25个汉字或50个英文字符
			</td>
		</tr>
		<tr height=20> 
			<td width=20% class=tablebody1><b>标题颜色</b></td>
			<td width=80% class=tablebody1><input type=hidden name=ChenTopicColor> <%
				if not CanTopicColor then
					%><font color="<%=forum_body(8)%>">您没有权限修改帖子主题颜色</font> <%
				else
					%><SCRIPT language=JavaScript src="inc/z_color.js"></SCRIPT>
					<SCRIPT language=JavaScript>
					var height0 = 10;		// define the height of the color bar
					var pas0 = 36;			// define the number of color in the color bar
					var width0=2;				//define the width of the color bar here (for forum with little width put 1)
					</SCRIPT>
					<TABLE id=HeadColorPanel cellSpacing=0 cellPadding=0 align=left border=0 >
					<TBODY>
						<TR><TD id=ColorHead onclick="this.bgColor='';MM_findObj('ChenTopicColor').value=''" vAlign=center align=middle><SCRIPT language=JavaScript>document.write('<IMG height='+height0+' width=20 border=1></TD>');</SCRIPT></TD><SCRIPT language=JavaScript>rgb1(pas0,width0,height0)</SCRIPT></TR>
					</TBODY>
					</TABLE><%
				end if
			%></td>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>投票项目 </b> <br><br><li>每行一个投票项目，最多<b><%=Board_Setting(32)%></b>个</li><li>超过自动作废，空行自动过滤</li><br><br><input type=radio name=votetype value=0 checked>单选投票<br><input type=radio name=votetype value=1>多选投票</font></td>
			<td width=80% class=tablebody2><textarea name=vote cols=103 rows=8></textarea></td>
		</tr>
		<tr>
			<td width=20% valign=top class=tablebody1><b>当前心情</b><br><li>将放在帖子的前面</td>
			<td width=80% class=tablebody1><%
				for i=0 to Forum_PostFaceNum-1
					%><input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;<%
					if i>0 and ((i+1) mod 9=0) then response.write "<br>"
				next
			%></td>
		</tr>
		<%if Cint(GroupSetting(7))=1 or Cint(GroupSetting(7))=2 then%>
			<tr>
				<td width=20% valign=middle class=tablebody2><b>文件上传</b>&nbsp;<select name="filetype" size=1><option value="">允许类型</option><%
					dim iupload
					iupload=split(Board_Setting(19),"|")
					for i=0 to ubound(iupload)
						response.write "<option value="&iupload(i)&">"&iupload(i)&"</option>"
					next
				%></select></td>
				<td width=80% class=tablebody2><iframe name="ad" frameborder=0 width=100% height=40 scrolling=no src=saveannounce_upload.asp?boardid=<%=boardid%>></iframe></td>
			</tr>
		<%end if%>
		<tr>
			<td width=20% valign=top class=tablebody1><b>内容</b><br><!--#include file="z_ContentFlag.asp"--></td>
			<td width=80% class=tablebody1><%if Cint(Board_Setting(4))=1 then%><!--#include file="inc/getubb.asp"--><%end if%><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交帖子 class=FormClass onkeydown=ctlent()></textarea><!--#include file="z_speak.asp"--><%
				if board_setting(48)=1 then
					%>&nbsp;&nbsp;短信引起注意&nbsp;<input name=msgNotify size=20 maxlength=20>&nbsp;<SELECT name=font onchange="document.frmAnnounce.msgNotify.value=(this.options[this.selectedIndex].value);document.frmAnnounce.msgNotify.focus();"><OPTION selected value="">选择</OPTION><%
					set rs=server.createobject("adodb.recordset")
					sql="select F_friend from Friend where F_username='"&membername&"' and f_type=0 order by F_addtime desc"
					rs.open sql,conn,1,1
					do while not rs.eof
						%><OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION><%
						rs.movenext
					loop
					rs.close
					set rs=nothing%></SELECT><%
				end if
			%></td>
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
			<td valign=middle class=tablebody1><%
				if board_setting(44)=1 then
					%><input type=checkbox name=ArticleRandom value=yes <%if board_setting(45)=1 then%>checked<%end if%>>是否使用发帖机遇？<br><%
				else
					%><input type=hidden name=ArticleRandom value=no><%
				end if
				%><input type=checkbox name=signflag value=yes checked>是否显示您的签名？<%
				if board_setting(46)=1 then
					%><br><input type=checkbox name=msgflag value=yes <%if board_setting(47)=1 then%>checked<%end if%>>有回复时使用短信通知您？(<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(28)
					else
						response.write forum_user(29)
					end if
					%>元，您现在有现金<%=mymoney%>元</font>)<%
				else
					%><input type=hidden name=msgflag value=no><%
				end if
				%><br><input type=checkbox name=emailflag value=yes>有回复时使用邮件通知您？(<font color=<%=forum_body(8)%>>使用此功能将花去您的现金<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(30)
				else
					response.write forum_user(31)
				end if
			%>元，您现在有现金<%=mymoney%>元</font>)</td>
		</tr>
		<tr>
			<td valign=middle colspan=2 align=center class=tablebody2><input type=Submit value='发 表' name=Submit onclick="if(checkbbs(true)){return true;} return false;"> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value='清 除'></td>
		</tr>
	</table>
	</form>
	<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
		<input type=hidden name=title value=><input type=hidden name=body value=>
	</form>
	<%call activeonline()
end sub%>