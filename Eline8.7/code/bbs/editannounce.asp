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
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
		end if
	end if
end if
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>VIP��Աר�ð���</font>����ȷ�����������Ƿ���ϡ�"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		dim sexshow
		if cint(Board_Setting(52))=1 then
			sexshow="Ů��"
		elseif cint(Board_Setting(52))=2 then
			sexshow="����"
		end if
		if chksexlogin(cint(Board_Setting(52)),membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>"&sexshow&"��̳����</font>����ȷ�������Ա��Ƿ���ϡ�"
			founderr=true
		end if
	end if
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	AnnounceID=request("id")
end if
if request("replyID")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("replyID")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	replyID=request("replyID")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
end if
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>����̳�Ѿ�������Ա�����˲���������"
	founderr=true
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
	end if
end if
stats=boardtype & "�༭����"
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
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ӧ�����ӡ�"
		Founderr=true
		exit sub
	else
		TotalUseTable=rs(0)
		ChenTopicColor=rs(1)
	end if
	sql="select b.username,b.topic,b.body,b.dateandtime,u.usergroupID,b.signflag,b.emailflag,b.speak,b.parentid from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.rootid="&AnnounceID&" and b.BoardID="&boardid&" and  b.AnnounceID="&replyID
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ӧ�����ӡ�"
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
			Errmsg=Errmsg+"<br>"+"<li>ϵͳ�༭����ʱ��Ϊ"&forum_setting(51)&"���ӣ��������������ӵ������Ѿ���"&Datediff("s",rs("dateandtime"),Now())\60&"�����ˡ�"
				founderr=true
				exit sub
			end if
		end if
		if rs("username")=membername then
			if Cint(GroupSetting(10))=0 then
				Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�༭�Լ����ӵ�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
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
				Errmsg=Errmsg+"<br>"+"<li>ͬ�ȼ��û������޸ġ�"
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
						Errmsg=Errmsg+"<br>"+"<li>ͬ�ȼ��û������޸ġ�"
						Founderr=true
						exit sub
					end if
				end if
			elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
				Errmsg=Errmsg+"<br>"+"<li>�����޸ĵȼ������ߵ��û������ӡ�"
				Founderr=true
				exit sub
			end if
			if not CanEditPost then
				Errmsg=Errmsg+"<br>"+"<li>��û���㹻��Ȩ�ޱ༭�����ӣ���͹���Ա��ϵ��"
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
	'�ж��û��Ƿ�������������ɫȨ��
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
			<th width=100% colspan=2>�༭����</th>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>�û���</b></td>
			<td width=80% class=tablebody2><input name=username value=<%=htmlencode(old_user)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>��û��ע�᣿</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody1><b>����</b></td>
			<td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=lostpass.asp>�������룿</a>�����༭����Ҫ����</td>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>�������</b>
				<SELECT name=font onchange=DoTitle2(this.options[this.selectedIndex].value)>
				<OPTION selected value="">ѡ����</OPTION> <OPTION value=[ԭ��]>[ԭ��]</OPTION> 
				<OPTION value=[ת��]>[ת��]</OPTION> <OPTION value=[��ˮ]>[��ˮ]</OPTION> 
				<OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION> 
				<OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> <OPTION value=[����]>[����]</OPTION> 
				<OPTION value=[ע��]>[ע��]</OPTION> <OPTION value=[��ͼ]>[��ͼ]</OPTION>
				<OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION>
				<OPTION value=[����]>[����]</OPTION>
				</SELECT>
			</td>
			<td width=80% class=tablebody2><input name=subject size=70 maxlength=80 value=<%=htmlencode(Topic)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>���ó��� 25 �����ֻ�50��Ӣ���ַ�</td>
		</tr>
		<tr height=20> 
			<td width=20% class=tablebody1><b>������ɫ</b></td>
			<td width=80% class=tablebody1><input type=hidden name=ChenTopicColor value="<%=ChenTopicColor%>"> <%
				if not CanTopicColor then
					%><font color="<%=forum_body(8)%>">��û��Ȩ���޸�����������ɫ������Ӳ�������</font> <%
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
			<td width=20% valign=top class=tablebody2><b>��ǰ����</b><br><li>���������ӵ�ǰ��</td>
			<td width=80% class=tablebody2><%
				for i=0 to Forum_PostFaceNum-1
					%><input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;<%
					if i>0 and ((i+1) mod 9=0) then response.write "<br>"
				next
			%></td>
		</tr>
		<tr>
			<td width=20% valign=top class=tablebody1><b>����</b><br><!--#include file="z_ContentFlag.asp"--></td>
			<td width=80% class=tablebody1><%if Cint(Board_Setting(4))=1 then%><!--#include file="inc/getubb.asp"--><%end if%><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL  onkeydown=ctlent()><%=server.htmlencode(content)%></textarea><!--#include file="z_speak.asp"--></td>
		</tr>
		<tr>
			<td class=tablebody2 valign=top colspan=2 style="table-layout:fixed; word-break:break-all"><b>�������ͼ�����������м�����Ӧ�ı���</B><br>&nbsp;<%
				dim ii
				for i=0 to 14
					if len(i)=1 then ii="0" & i else ii=i
					response.write "<img src="""&Forum_info(10)&Forum_emot(i)&""" border=0 onclick=""insertsmilie('[em"&ii&"]')"" style=""CURSOR: hand"">&nbsp;"
				next
			%></td>
		</tr>
		<tr>
			<td valign=top class=tablebody1><b>ѡ��</b></td>
			<td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes <%if signflag then%>checked<%end if%>>�Ƿ���ʾ����ǩ����<%
				if board_setting(46)=1 then
					%><br><input type=checkbox name=msgflag value=yes <%if msgflag=1 then%>checked<%end if%>>�лظ�ʱʹ�ö���֪ͨ����(<font color=<%=forum_body(8)%>>ʹ�ô˹��ܽ���ȥ�����ֽ�<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(28)
					else
						response.write forum_user(29)
					end if
					%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font>)<%
				else
					%><input type=hidden name=msgflag value=no><%
				end if
				%><br><input type=checkbox name=emailflag value=yes <%if mailflag=1 then%>checked<%end if%>>�лظ�ʱʹ���ʼ�֪ͨ����(<font color=<%=forum_body(8)%>>ʹ�ô˹��ܽ���ȥ�����ֽ�<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(30)
				else
					response.write forum_user(31)
				end if
			%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font>)</td>
		</tr>
		<tr>
			<td class=tablebody2 valign=middle colspan=2 align=center><input type=Submit value="�� ��" name=Submit onclick="if(checkbbs(<%if isTopic then%>true<%else%>false<%end if%>)){return true;} return false;"> &nbsp; <input type=button value='Ԥ ��' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value="�� ��"></td>
		</tr>
	</table>
	</form>
	<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
		<input type=hidden name=title value=><input type=hidden name=body value=>
	</form>
<%end sub%>