<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_fastpost_conn.asp"-->
<%
dim msg
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
stats="���ٷ����б�"
call nav()
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
%>
<!--#include file="z_controlpanel.asp"-->
<br>
<%
	select case request("action")
	case "info"
		call info()
	case "new"
		call addF()
	case "edit"
		call editF()
	case "newsave"
		call newsave()
	case "editsave"
		call editsave()
	case "ɾ��"
		call DelF()
	case "��տ��ٷ���"
		call AllDelF()
	case else
		call info()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
connfastpost.close
set connfastpost=nothing
call footer()

sub info()%>
	<form action="z_fastpost.asp" method=post name=inbox>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
		<tr>
			<th valign=middle width="25%">��ǩ����</td>
			<th valign=middle width="*">��ǩ����</td>
			<th valign=middle width="5%">����</td>
		</tr>
		<%set rs=server.createobject("adodb.recordset")
		sql="select * from FastPost where username='"&trim(membername)&"'"
		rs.open sql,connfastpost,1,1
		if rs.eof and rs.bof then
			%><tr>
				<td class=tablebody1 align=center valign=middle colspan=6>���Ŀ��ٷ����б���û���κ����ݡ�</td>
			</tr>
		<%else
			do while not rs.eof%>
				<tr>
					<td align=center valign=middle class=tablebody1><a href="z_fastpost.asp?action=edit&id=<%=rs("ID")%>"><%=htmlencode(rs("LabelName"))%></a></td>
					<td align=left valign=middle class=tablebody1><blockquote><%=htmlencode(rs("labelcontent"))%></blockquote></td>
					<td align=center class=tablebody1><input type=checkbox name=id value=<%=rs("id")%>></td>
				</tr>
				<%rs.movenext
			loop
		end if
		rs.close
		set rs=nothing%>
		<tr> 
			<td align=right valign=middle colspan=3 class=tablebody2><input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ѡ��������ʾ��¼&nbsp;<input type=button name=action onclick="location.href='z_fastpost.asp?action=new'" value="��ӿ��ٷ���">&nbsp;<input type=submit name=action onclick="{if(confirm('ȷ��ɾ��ѡ���ļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="ɾ��">&nbsp;<input type=submit name=action onclick="{if(confirm('ȷ��������еļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="��տ��ٷ���"></td>
		</tr>
	</table>
	</form>
<%end sub

sub delF()
	dim delid

	delid=replace(request.form("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	if delid="" or isnull(delid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id in ("&delid&")")(0)<1 then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		founderr=true
		exit sub
	else
		connfastpost.execute("delete from Fastpost where username='"&trim(membername)&"' and id in ("&delid&")")
		sucmsg=sucmsg+"<br>"+"<li>���Ѿ�ɾ��ѡ���Ŀ��ٷ�����"
		call dvbbs_suc()
	end if
end sub

sub AllDelF()
	connfastpost.execute("delete from Fastpost where username='"&trim(membername)&"'")
	sucmsg=sucmsg+"<br>"+"<li>���Ѿ�ɾ�������п��ٷ����б�"
	call dvbbs_suc()
end sub

sub addF()%>
	<script src="inc/ubbcode.js"></script>
	<form method=POST name=frmAnnounce action="?action=newsave">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<TBODY>
		<tr><th width=100% colspan="2" height=23 id=tabletitlelink>�û����ٷ������� | <a href=z_fastpost.asp>����������ҳ</a></th></tr>
		<tr>
			<td width="20%" valign=middle class=tablebody1 align=left><b>�û����ٷ�������˵����</b><br><br><li>�������ڷ���ʱֱ�ӵ�<br>&nbsp;&nbsp;&nbsp;�ô����������ٷ���<br><br><li>��������������������</td>
			<td width="80%" valign=middle class=tablebody1 align=left>��ǩ����&nbsp;<input name=subject size=50 maxlength=20>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><strong>*</strong></font>���ó��� 10 �����ֻ�20��Ӣ���ַ�<br><!--#include file="inc/getubb.asp"--><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title="��д�������õĶ��֧�ָ�����������ʽ��������ʹ��Ctrl+Enterֱ���ύ" class=FormClass onkeydown=ctlent()></textarea></td>
		</tr>
		<tr><td class=tablebody2 valign=middle align=center colspan="2" height="36"><input type=Submit value="�� ��" name=Submit>&nbsp;&nbsp;<input type=reset name=Submit2 value="�� ��"></td></tr>
	</TBODY>
	</table>
	</form>
<%end sub

sub editF()
	dim editid

	editid=replace(request("id"),"'","")
	if editid="" or isnull(editid) then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id="&editid&"")(0)<1 then
		Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
		founderr=true
		exit sub
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from FastPost where id="&editid&""
	rs.open sql,connfastpost,1,1
	%><script src="inc/ubbcode.js"></script>
	<form method=POST name=frmAnnounce action="?action=editsave&id=<%=rs("ID")%>">
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
	<TBODY>
		<tr><th width=100% colspan="2" height=23 id=tabletitlelink>�û����ٷ������� | <a href=z_fastpost.asp>����������ҳ</a></th></tr>
		<tr>
			<td width="20%" valign=middle class=tablebody1 align=left><b>�û����ٷ�������˵����</b><br><br><li>�������ڷ���ʱֱ�ӵ�<br>&nbsp;&nbsp;&nbsp;�ô����������ٷ���<br><br><li>��������������������</td>
			<td width="80%" valign=middle class=tablebody1 align=left>��ǩ����&nbsp;<input name=subject size=50 maxlength=20 value=<%=rs("labelname")%>>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><strong>*</strong></font>���ó��� 10 �����ֻ�20��Ӣ���ַ�<br><!--#include file="inc/getubb.asp"--><textarea class=smallarea cols=103 name=Content rows=13 wrap=VIRTUAL title="��д�������õĶ��֧�ָ�����������ʽ��������ʹ��Ctrl+Enterֱ���ύ" class=FormClass onkeydown=ctlent()><%=rs("labelcontent")%></textarea></td>
		</tr>
		<tr><td class=tablebody2 valign=middle align=center colspan="2" height="36"><input type=Submit value="�� ��" name=Submit>&nbsp;&nbsp;<input type=reset name=Submit2 value="�� ��"></td></tr>
	</TBODY>
	</table>
	</form>
	<%rs.close
	set rs=nothing
end sub

sub newsave()
	dim labelname
	dim labelcontent
	
	labelname=replace(request("subject"),"'","")
	labelcontent=replace(request("content"),"'","")
	if labelname="" then
		errmsg=errmsg+"<li>��������д��ǩ�����˰ɡ�"
		founderr=true
		exit sub
	elseif labelContent="" then
		errmsg=errmsg+"<li>��������д��ǩ�����˰ɡ�"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and labelname='"&labelname&"'")(0)>0 then
		Errmsg=Errmsg+"<li>�ñ�ǩ�����Ѿ����ڣ������ٴ���ӡ�"
		founderr=true
		exit sub
	end if
	connfastpost.execute("insert into fastpost (username,labelname,labelcontent) values ('"&membername&"','"&labelname&"','"&labelcontent&"')")
	sucmsg=sucmsg+"<li>��ϲ�����Ѿ��ɹ�������˿��ٷ�����"
	call dvbbs_suc()
end sub

sub editsave()
	dim editid
	dim labelname
	dim labelcontent
	
	editid=replace(request("id"),"'","")
	labelname=replace(request("subject"),"'","")
	labelcontent=replace(request("content"),"'","")
	if editid="" or isnull(editid) then
		Errmsg=Errmsg+"<li>��ѡ����ز�����"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and id="&editid&"")(0)<1 then
		Errmsg=Errmsg+"<li>��ѡ����ز�����"
		founderr=true
		exit sub
	elseif labelname="" then
		errmsg=errmsg+"<li>��������д��ǩ�����˰ɡ�"
		founderr=true
		exit sub
	elseif labelContent="" then
		errmsg=errmsg+"<li>��������д��ǩ�����˰ɡ�"
		founderr=true
		exit sub
	elseif connfastpost.execute("select count(*) from fastpost where username='"&trim(membername)&"' and labelname='"&labelname&"' and id<>"&editid&"")(0)>0 then
		Errmsg=Errmsg+"<li>�����޸�Ϊ�Ѿ����ڵı�ǩ���ơ�"
		founderr=true
		exit sub
	end if
	connfastpost.execute("update fastpost set labelname='"&labelname&"',labelcontent='"&labelcontent&"' where id="&editid&"")
	sucmsg=sucmsg+"<li>��ϲ�����Ѿ��ɹ����޸��˿��ٷ�����"
	call dvbbs_suc()
end sub
%>