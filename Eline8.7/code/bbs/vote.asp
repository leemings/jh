<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="z_fastpost_conn.asp"-->
<%if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats=boardtype&"����ͶƱ"
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
	dim CanTopicColor
	if (Cint(Board_Setting(43))=1 or Cint(Board_Setting(0))=1) and not (master or superboardmaster or boardmaster) then
		Errmsg=Errmsg+"<br>"+"<li>����̳�Ѿ�������Ա�����˲�����������"
		founderr=true
		exit sub
	end if
	if cint(Board_Setting(1))=1 then
		if Cint(GroupSetting(37))=0 then
			Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
			founderr=true
			exit sub
		end if
	end if
	if cint(Board_Setting(51))=1 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
			exit sub
		else
			if chkviplogin(membername)=false then
				Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>VIP��Աר�ð���</font>����ȷ�����������Ƿ���ϡ�"
				founderr=true
				exit sub
			end if
		end if
	end if
	if cint(Board_Setting(52))<>0 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
			exit sub
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
				exit sub
			end if
		end if
	end if
	if Cint(GroupSetting(8))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳����ͶƱ��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
	end if
	if cint(lockboard)=2 then
		if not founduser then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
			exit sub
		else
			if chkboardlogin(boardid,membername)=false then
				Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
				founderr=true
				exit sub
			end if
		end if
	end if
	'�ж��û��Ƿ�������������ɫȨ��
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
			<th width=100% colspan=2 align=left>&nbsp;&nbsp;������ͶƱ<%if isaudit=1 then%>���������������Ӷ�Ҫ��������Ա��˷��ɷ�����<%end if%>  </th>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>�û���</b></td>
			<td width=80% class=tablebody2><input name=username value=<%=membername%> class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>��û��ע�᣿</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody1><b>����</b></td>
			<td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> class=FormClass><font color=<%=Forum_body(8)%>>&nbsp;&nbsp;<b>*</b></font><a href=lostpass.asp>�������룿</a></td>
		</tr>
		<tr>
			<td width=20% class=tablebody2><b>�������</b>&nbsp;<SELECT name=font onchange=DoTitle2(this.options[this.selectedIndex].value)>
				<OPTION selected value="">ѡ����</OPTION> <OPTION value=[ԭ��]>[ԭ��]</OPTION> 
				<OPTION value=[ת��]>[ת��]</OPTION> <OPTION value=[��ˮ]>[��ˮ]</OPTION> 
				<OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION> 
				<OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> <OPTION value=[����]>[����]</OPTION> 
				<OPTION value=[ע��]>[ע��]</OPTION> <OPTION value=[��ͼ]>[��ͼ]</OPTION>
				<OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION>
				<OPTION value=[����]>[����]</OPTION></SELECT>
			</td>
			<td width=80% class=tablebody2><input name=subject size=50 maxlength=80 class=FormClass>&nbsp;<select name=votetimeout size=1>
				<option value=0>����ʱ��</option>
				<option value=0>��������</option>
				<option value=1>һ��</option>
				<option value=3>����</option>
				<option value=7>һ��</option>
				<option value=15>����</option>
				<option value=30>һ��</option>
				<option value=90>����</option>
				<option value=180>����</option>
				</select>&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>���ó���25�����ֻ�50��Ӣ���ַ�
			</td>
		</tr>
		<tr height=20> 
			<td width=20% class=tablebody1><b>������ɫ</b></td>
			<td width=80% class=tablebody1><input type=hidden name=ChenTopicColor> <%
				if not CanTopicColor then
					%><font color="<%=forum_body(8)%>">��û��Ȩ���޸�����������ɫ</font> <%
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
			<td width=20% class=tablebody2><b>ͶƱ��Ŀ </b> <br><br><li>ÿ��һ��ͶƱ��Ŀ�����<b><%=Board_Setting(32)%></b>��</li><li>�����Զ����ϣ������Զ�����</li><br><br><input type=radio name=votetype value=0 checked>��ѡͶƱ<br><input type=radio name=votetype value=1>��ѡͶƱ</font></td>
			<td width=80% class=tablebody2><textarea name=vote cols=103 rows=8></textarea></td>
		</tr>
		<tr>
			<td width=20% valign=top class=tablebody1><b>��ǰ����</b><br><li>���������ӵ�ǰ��</td>
			<td width=80% class=tablebody1><%
				for i=0 to Forum_PostFaceNum-1
					%><input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;<%
					if i>0 and ((i+1) mod 9=0) then response.write "<br>"
				next
			%></td>
		</tr>
		<%if Cint(GroupSetting(7))=1 or Cint(GroupSetting(7))=2 then%>
			<tr>
				<td width=20% valign=middle class=tablebody2><b>�ļ��ϴ�</b>&nbsp;<select name="filetype" size=1><option value="">��������</option><%
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
			<td width=20% valign=top class=tablebody1><b>����</b><br><!--#include file="z_ContentFlag.asp"--></td>
			<td width=80% class=tablebody1><%if Cint(Board_Setting(4))=1 then%><!--#include file="inc/getubb.asp"--><%end if%><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title=����ʹ��Ctrl+Enterֱ���ύ���� class=FormClass onkeydown=ctlent()></textarea><!--#include file="z_speak.asp"--><%
				if board_setting(48)=1 then
					%>&nbsp;&nbsp;��������ע��&nbsp;<input name=msgNotify size=20 maxlength=20>&nbsp;<SELECT name=font onchange="document.frmAnnounce.msgNotify.value=(this.options[this.selectedIndex].value);document.frmAnnounce.msgNotify.focus();"><OPTION selected value="">ѡ��</OPTION><%
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
			<td valign=middle class=tablebody1><%
				if board_setting(44)=1 then
					%><input type=checkbox name=ArticleRandom value=yes <%if board_setting(45)=1 then%>checked<%end if%>>�Ƿ�ʹ�÷���������<br><%
				else
					%><input type=hidden name=ArticleRandom value=no><%
				end if
				%><input type=checkbox name=signflag value=yes checked>�Ƿ���ʾ����ǩ����<%
				if board_setting(46)=1 then
					%><br><input type=checkbox name=msgflag value=yes <%if board_setting(47)=1 then%>checked<%end if%>>�лظ�ʱʹ�ö���֪ͨ����(<font color=<%=forum_body(8)%>>ʹ�ô˹��ܽ���ȥ�����ֽ�<%
					if isnull(myvip) or myvip<>1 then
						response.write forum_user(28)
					else
						response.write forum_user(29)
					end if
					%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font>)<%
				else
					%><input type=hidden name=msgflag value=no><%
				end if
				%><br><input type=checkbox name=emailflag value=yes>�лظ�ʱʹ���ʼ�֪ͨ����(<font color=<%=forum_body(8)%>>ʹ�ô˹��ܽ���ȥ�����ֽ�<%
				if isnull(myvip) or myvip<>1 then
					response.write forum_user(30)
				else
					response.write forum_user(31)
				end if
			%>Ԫ�����������ֽ�<%=mymoney%>Ԫ</font>)</td>
		</tr>
		<tr>
			<td valign=middle colspan=2 align=center class=tablebody2><input type=Submit value='�� ��' name=Submit onclick="if(checkbbs(true)){return true;} return false;"> &nbsp; <input type=button value='Ԥ ��' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value='�� ��'></td>
		</tr>
	</table>
	</form>
	<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
		<input type=hidden name=title value=><input type=hidden name=body value=>
	</form>
	<%call activeonline()
end sub%>