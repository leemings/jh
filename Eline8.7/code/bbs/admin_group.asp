<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="25"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
if request("action")="save" then
call savegroup()
elseif request("action")="savedit" then
call savedit()
elseif request("action")="del" then
call del()
elseif request("action")="group" then
call gradeinfo()
elseif request("action")="addgroup" then
call addgroup()
elseif request("action")="editgroup" then
call editgroup()
elseif request("action")="delgroup" then
call delgroup()
else
call usergroup()
end if
end sub

sub usergroup()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="4" >�û������&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF>[����û���]</font></a></th>
</tr>
<tr align=center>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>�û���</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>�û�����</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>�༭</B></td>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>�г��û�</B></td>
</tr>
<%
dim trs
set rs=conn.execute("select * from usergroups order by usergroupID")
do while not rs.eof
set trs=conn.execute("select count(*) from [user] where UserGroupID="&rs("UserGroupID"))
%>
<tr align=center>
<td height="23" width="30%" class="Forumrow"><%=rs("title")%></td>
<td height="23" width="20%" class="Forumrow"><%=trs(0)%></td>
<td height="23" width="20%" class="Forumrow"><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a><%if rs("UserGroupID")>8 then%> | <a href="?action=delgroup&groupid=<%=rs("UserGroupID")%>">ɾ��</a><%end if%></td>
<td height="23" width="30%" class="Forumrow"><a href="admin_user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г������û�</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr><td colspan=4 height=25 class="Forumrow"><B>˵��</B>��<BR>�����������������ø����û�������̳�е�Ĭ��Ȩ�ޣ���̳Ĭ���û��鲻��ɾ���ͱ༭�û�����<BR>�ڿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�������ĳ�����û����й������������û��б�<BR>�ۿ���ɾ���ͱ༭����ӵ��û���<BR>����������õ��ȼ�����������ĵȼ�������˵���뿴��ӵȼ�ҳ�棬����������Ӧ�ȼ���������û��ĵȼ�Ϊ�û��������ȼ�ͼ��Ϊ��̳��͵ȼ�ͼ��<BR>��<B>���ɾ���û��飬����û������������û����Զ�ת��ע���û��飬ͬʱɾ���ڵȼ��к͸��û�������ĵȼ��������¸��û������������û��ĵȼ�Ϊע���û��鰴�����¼���ĵȼ�</B></td></tr>
</table>
<%
end sub

sub delgroup()
if not isnumeric(request("groupid")) then
response.write "����Ĳ�����"
exit sub
end if
'�����û��ȼ�����
Server.ScriptTimeout=999999
dim UserGrade
set rs=conn.execute("select userid,article from [user] where UserGroupID="&request("groupid"))
do while not rs.eof
	set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where UserGroupID=4 and Minarticle<="&rs(1)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
	conn.execute("update [user] set userclass='"&rs(0)&"',titlepic='"&rs(1)&"' where userid="&rs(0))
rs.movenext
loop
set rs=nothing
'�����û�������
conn.execute("update [user] set UserGroupID=4 where UserGroupID="&request("groupid"))
'ɾ�������ȼ�
conn.execute("delete from usertitle where UserGroupID="&request("groupid"))
'ɾ���û���
conn.execute("delete from usergroups where UserGroupID="&request("groupid"))
response.write "ɾ���ɹ���"
end sub

sub editgroup()
if not isnumeric(request("groupid")) then
response.write "����Ĳ�����"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "�������û������ƣ�"
	exit sub
	end if
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & request.form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from usergroups where usergroupid="&request("groupid")
	rs.open sql,conn,1,3
	rs("title")=request.form("title")
	rs("GroupSetting")=GroupSetting
	rs.update
	rs.close
	set rs=nothing
	response.write "�޸ĳɹ���"
else
Dim reGroupSetting
set rs=conn.execute("select * from usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "δ�ҵ����û��飡"
exit sub
end if
reGroupSetting=split(rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup">
<input type=hidden name="groupid" value="<%=request("groupid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><td colspan=2 height=25 class="Forumrow"><B>˵��</B>��<BR>�����������������ø����û�������̳�е�Ĭ��Ȩ�ޣ���̳Ĭ���û��鲻��ɾ���ͱ༭�û�����<BR>�ڿ���ɾ���ͱ༭����ӵ��û���</td></tr>
<tr> 
<th height="23" colspan="2">�༭�û��飺<%=rs("title")%></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�û�������</td>
<td height="23" width="40%" class=Forumrow><input size=35 name="title" type=text value="<%=rs("title")%>"  <%if Cint(request("GroupID"))<9 then%>disabled<%end if%>></td>
</tr>
<%if Cint(request("GroupID"))<9 then%>
<input name="title" type=hidden value="<%=rs("title")%>">
<%end if%>
<tr> 
<th height="23" colspan="2"  align=left>�����鿴Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳</td>
<td height="23" width="40%" class=Forumrow>��<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;��<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2"  align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���������</td>
<td height="23" width="40%" class=Forumrow>��<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;��<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Իظ������˵�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������������Ǯ
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ϴ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;��<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>>
&nbsp;���������ϴ�<input name="canupload" type=radio value="2" <%if reGroupSetting(7)=2 then%>checked<%end if%>>&nbsp;�ظ������ϴ�<input name="canupload" type=radio value="3" <%if reGroupSetting(7)=3 then%>checked<%end if%>>
</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="ba2" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�ϴ��ļ���С����
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�����ͶƱ</td>
<td height="23" width="40%" class=Forumrow>��<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ���ͶƱ</td>
<td height="23" width="40%" class=Forumrow>��<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;��<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���С�ֱ�</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansmallpaper" type=radio value="1"  <%if reGroupSetting(17)=1 then%>checked<%end if%>>&nbsp;��<input name="cansmallpaper" type=radio value="0" <%if reGroupSetting(17)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����С�ֱ������Ǯ</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="<%=reGroupSetting(46)%>" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2"  align=left>����<b>����/����༭Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;��<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;��<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;��<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;��<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޸ĸ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳�¼�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canvieweven" type=radio value="1"  <%if reGroupSetting(39)=1 then%>checked<%end if%>>&nbsp;��<input name="canvieweven" type=radio value="0" <%if reGroupSetting(39)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������̳չ����Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="ba1" type=radio value="1"  <%if reGroupSetting(49)=1 then%>checked<%end if%>>&nbsp;��<input name="ba1" type=radio value="0" <%if reGroupSetting(49)=0 then%>checked<%end if%>></td>
</tr>
<input type=hidden value=0 name="ba4">
<input type=hidden value=0 name="ba5">
<input type=hidden value=0 name="ba6">
<input type=hidden value=0 name="ba7">
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="candeltopic" type=radio value="1" <%if reGroupSetting(18)=1 then%>checked<%end if%>>&nbsp;��<input name="candeltopic" type=radio value="0"  <%if reGroupSetting(18)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmovetopic" type=radio value="1" <%if reGroupSetting(19)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovetopic" type=radio value="0"  <%if reGroupSetting(19)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canclosetopic" type=radio value="1" <%if reGroupSetting(20)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosetopic" type=radio value="0"  <%if reGroupSetting(20)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ̶�/����̶�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cantoptopic" type=radio value="1" <%if reGroupSetting(21)=1 then%>checked<%end if%>>&nbsp;��<input name="cantoptopic" type=radio value="0"  <%if reGroupSetting(21)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ��������̶ܹ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusesign" type=radio value="1"  <%if reGroupSetting(38)=1 then%>checked<%end if%>>&nbsp;��<input name="canusesign" type=radio value="0" <%if reGroupSetting(38)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canawardtopic" type=radio value="1" <%if reGroupSetting(22)=1 then%>checked<%end if%>>&nbsp;��<input name="canawardtopic" type=radio value="0"  <%if reGroupSetting(22)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ��û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canaward" type=radio value="1" <%if reGroupSetting(43)=1 then%>checked<%end if%>>&nbsp;��<input name="canaward" type=radio value="0"  <%if reGroupSetting(43)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmodifytopic" type=radio value="1" <%if reGroupSetting(23)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodifytopic" type=radio value="0" <%if reGroupSetting(23)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Լ���/�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbesttopic" type=radio value="1" <%if reGroupSetting(24)=1 then%>checked<%end if%>>&nbsp;��<input name="canbesttopic" type=radio value="0"  <%if reGroupSetting(24)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAnnounce" type=radio value="1" <%if reGroupSetting(25)=1 then%>checked<%end if%>>&nbsp;��<input name="canAnnounce" type=radio value="0"  <%if reGroupSetting(25)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminAnnounce" type=radio value="1" <%if reGroupSetting(26)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminAnnounce" type=radio value="0"  <%if reGroupSetting(26)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ���С�ֱ�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminPaper" type=radio value="1" <%if reGroupSetting(27)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminPaper" type=radio value="0"  <%if reGroupSetting(27)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������/����/��������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminUser" type=radio value="1" <%if reGroupSetting(28)=1 then%>checked<%end if%>>&nbsp;��<input name="canAdminUser" type=radio value="0"  <%if reGroupSetting(28)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���û�1��10������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canDelUserTopic" type=radio value="1" <%if reGroupSetting(29)=1 then%>checked<%end if%>>&nbsp;��<input name="canDelUserTopic" type=radio value="0"  <%if reGroupSetting(29)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴����IP����Դ
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewip" type=radio value="1" <%if reGroupSetting(30)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewip" type=radio value="0"  <%if reGroupSetting(30)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޶�IP����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canadminip" type=radio value="1" <%if reGroupSetting(31)=1 then%>checked<%end if%>>&nbsp;��<input name="canadminip" type=radio value="0"  <%if reGroupSetting(31)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ����û�Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="adminpermission" type=radio value="1" <%if reGroupSetting(42)=1 then%>checked<%end if%>>&nbsp;��<input name="adminpermission" type=radio value="0"  <%if reGroupSetting(42)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������ɾ�����ӣ�ǰ̨��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbatchtopic" type=radio value="1" <%if reGroupSetting(45)=1 then%>checked<%end if%>>&nbsp;��<input name="canbatchtopic" type=radio value="0"  <%if reGroupSetting(45)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ���������ӵ�Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusetitle" type=radio value="1" <%if reGroupSetting(36)=1 then%>checked<%end if%>>&nbsp;��<input name="canusetitle" type=radio value="0" <%if reGroupSetting(36)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ��н���������̳��Ȩ��<BR>���û����п�����ɽ�������������̳<BR>����̳Ȩ�޹��������ÿɽ������õİ���<BR>���û�Ȩ�޹��������ÿɽ������õİ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canuseface" type=radio value="1"  <%if reGroupSetting(37)=1 then%>checked<%end if%>>&nbsp;��<input name="canuseface" type=radio value="0" <%if reGroupSetting(37)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����̳�ļ�����Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canadminfile" type=radio value="1" <%if reGroupSetting(48)=1 then%>checked<%end if%>>&nbsp;��<input name="canadminfile" type=radio value="0" <%if reGroupSetting(48)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>������������ɫȨ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="ba3" type=radio value="1" <%if reGroupSetting(51)=1 then%>checked<%end if%>>&nbsp;��<input name="ba3" type=radio value="0" <%if reGroupSetting(51)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է��Ͷ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansendsms" type=radio value="1"  <%if reGroupSetting(32)=1 then%>checked<%end if%>>&nbsp;��<input name="cansendsms" type=radio value="0" <%if reGroupSetting(32)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��෢���û�
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������ݴ�С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
</tr>
<tr bgcolor=<%=Forum_body(2)%>> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow><input type="submit" name="submit" value="�� ��"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub addgroup()
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "�������û������ƣ�"
	exit sub
	end if
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & request.form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	'response.write GroupSetting
	'response.end
	set rs=conn.execute("select * from usertitle where usertitle='"&request.form("title")&"'")
	if not (rs.eof and rs.bof) then
		Errmsg="<br><li>�õȼ������Ѿ����ڡ�"
		call dvbbs_error()
		exit sub
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from usergroups where title='"&request.form("title")&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("title")=request.form("title")
		rs("GroupSetting")=GroupSetting
		rs.update
	else
		Errmsg="<br><li>���û��������Ѿ����ڡ�"
		call dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	dim maxid,maxuserclass,toptitlepic
	set rs=conn.execute("select top 1 * from usergroups order by usergroupid desc")
	maxid=rs("usergroupid")
	set rs=conn.execute("select max(userclass) from usertitle")
	maxuserclass=rs(0)+1
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from usertitle where usertitle='"&request.form("title")&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("userclass")=maxuserclass
		rs("usertitle")=request.form("title")
		rs("minarticle")=-1
		rs("titlepic")="level10.gif"
		rs("usergroupid")=maxid
		rs.update
	else
		Errmsg="<br><li>�õȼ������Ѿ����ڡ�"
		call dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	response.write "��ӳɹ�������ͬʱ�Ը�����������̳�ȼ����µȼ�������Ĭ�ϵ�ͼƬ���ã������Ե��ȼ������н����޸ģ�"
else
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=addgroup">
<tr><td colspan=2 height=25 class="Forumrow"><B>˵��</B>��<BR>�ٿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�������ĳ�����û����й������������û��б�<BR>�ڿ���ɾ���ͱ༭����ӵ��û���<BR>��<B>��������õ��ȼ�����������ĵȼ�</B>������˵���뿴��ӵȼ�ҳ�棬����������Ӧ�ȼ���������û��ĵȼ�Ϊ�û��������ȼ�ͼ��Ϊ��̳��͵ȼ�ͼ��</td></tr>
<tr> 
<th height="23" colspan="2" >����µ��û���</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�û�������</td>
<td height="23" width="40%" class=Forumrow><input size=35 name="title" type=text></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>�����鿴Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳</td>
<td height="23" width="40%" class=Forumrow>��<input name="canview" type=radio value="1" checked>&nbsp;��<input name="canview" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewuserinfo" type=radio value="1" checked>&nbsp;��<input name="canviewuserinfo" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewpost" type=radio value="1" checked>&nbsp;��<input name="canviewpost" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewbest" type=radio value="1" checked>&nbsp;��<input name="canviewbest" type=radio value="0"></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���������</td>
<td height="23" width="40%" class=Forumrow>��<input name="cannewpost" type=radio value="1" checked>&nbsp;��<input name="cannewpost" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canreplymytopic" type=radio value="1" checked>&nbsp;��<input name="canreplymytopic" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Իظ������˵�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canreplytopic" type=radio value="1" checked>&nbsp;��<input name="canreplytopic" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canpostagree" type=radio value="1" checked>&nbsp;��<input name="canpostagree" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������������Ǯ
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="5"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ϴ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canupload" type=radio value="1" checked>&nbsp;��<input name="canupload" type=radio value="0">&nbsp;���������ϴ�<input name="canupload" type=radio value="2">&nbsp;�ظ������ϴ�<input name="canupload" type=radio value="3"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="canuploadnum" type=text size=4 value="3"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="ba2" type=text size=4 value="30"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�ϴ��ļ���С����
</td>
<td height="23" width="40%" class=Forumrow><input name="MaxUploadSize" type=text size=4 value="100"> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�����ͶƱ</td>
<td height="23" width="40%" class=Forumrow>��<input name="canpostvote" type=radio value="1" checked>&nbsp;��<input name="canpostvote" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ���ͶƱ</td>
<td height="23" width="40%" class=Forumrow>��<input name="canvote" type=radio value="1" checked>&nbsp;��<input name="canvote" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���С�ֱ�</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansmallpaper" type=radio value="1" checked>&nbsp;��<input name="cansmallpaper" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����С�ֱ������Ǯ</td>
<td height="23" width="40%" class=Forumrow><input name="smallpapermoney" type=text value="50" size=4></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������/����༭Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="caneditmytopic" type=radio value="1">&nbsp;��<input name="caneditmytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="candelmytopic" type=radio value="1">&nbsp;��<input name="candelmytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmovemytopic" type=radio value="1">&nbsp;��<input name="canmovemytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canclosemytopic" type=radio value="1">&nbsp;��<input name="canclosemytopic" type=radio value="0" checked></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansearch" type=radio value="1" checked>&nbsp;��<input name="cansearch" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmailtopic" type=radio value="1">&nbsp;��<input name="canmailtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޸ĸ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmodify" type=radio value="1" checked>&nbsp;��<input name="canmodify" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳�¼�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canvieweven" type=radio value="1" checked>&nbsp;��<input name="canvieweven" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������̳չ����Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="ba1" type=radio value="1">&nbsp;��<input name="ba1" type=radio value="0" checked></td>
</tr>
<input type=hidden value=0 name="ba4">
<input type=hidden value=0 name="ba5">
<input type=hidden value=0 name="ba6">
<input type=hidden value=0 name="ba7">
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="candeltopic" type=radio value="1">&nbsp;��<input name="candeltopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmovetopic" type=radio value="1">&nbsp;��<input name="canmovetopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canclosetopic" type=radio value="1">&nbsp;��<input name="canclosetopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ̶�/����̶�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cantoptopic" type=radio value="1">&nbsp;��<input name="cantoptopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ��������̶ܹ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusesign" type=radio value="1">&nbsp;��<input name="canusesign" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canawardtopic" type=radio value="1">&nbsp;��<input name="canawardtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ��û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canaward" type=radio value="1">&nbsp;��<input name="canaward" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canmodifytopic" type=radio value="1">&nbsp;��<input name="canmodifytopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Լ���/�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbesttopic" type=radio value="1">&nbsp;��<input name="canbesttopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAnnounce" type=radio value="1">&nbsp;��<input name="canAnnounce" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminAnnounce" type=radio value="1">&nbsp;��<input name="canAdminAnnounce" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ���С�ֱ�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminPaper" type=radio value="1">&nbsp;��<input name="canAdminPaper" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������/����/��������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canAdminUser" type=radio value="1">&nbsp;��<input name="canAdminUser" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���û�1��10������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canDelUserTopic" type=radio value="1">&nbsp;��<input name="canDelUserTopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴����IP����Դ
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewip" type=radio value="1">&nbsp;��<input name="canviewip" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޶�IP����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canadminip" type=radio value="1">&nbsp;��<input name="canadminip" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ����û�Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="adminpermission" type=radio value="1">&nbsp;��<input name="adminpermission" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������ɾ�����ӣ�ǰ̨��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canbatchtopic" type=radio value="1">&nbsp;��<input name="canbatchtopic" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ���������ӵ�Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canusetitle" type=radio value="1" checked>&nbsp;��<input name="canusetitle" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ��н���������̳��Ȩ��<BR>���û����п�����ɽ�������������̳<BR>����̳Ȩ�޹��������ÿɽ������õİ���<BR>���û�Ȩ�޹��������ÿɽ������õİ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canuseface" type=radio value="1" checked>&nbsp;��<input name="canuseface" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����̳�ļ�����Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canadminfile" type=radio value="1">&nbsp;��<input name="canadminfile" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>������������ɫȨ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="ba3" type=radio value="1">&nbsp;��<input name="ba3" type=radio value="0" checked></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է��Ͷ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="cansendsms" type=radio value="1" checked>&nbsp;��<input name="cansendsms" type=radio value="0"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��෢���û�
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsendsms" size=5 type=text value="5"></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������ݴ�С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbody" size=5 type=text value="300"> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����С����
</td>
<td height="23" width="40%" class=Forumrow><input name="Maxsmsbox" size=5 type=text value="100"> KB</td>
</tr>
<tr> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow><input type="submit" name="submit" value="�� ��"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub gradeinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center bordercolor=<%=Forum_body(1)%>>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>�û����ɹ���</b>������������޸Ļ���ɾ����̳���ɡ�</font></td>
</tr>
<%if request("action")="edit" then%>
<form method="POST" action=admin_group.asp?action=savedit>
<%
	set rs=conn.execute("select * from GroupName where id="&request("id"))
%>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>�޸�����</b> | <a href=admin_group.asp><font color="<%=Forum_body(7)%>">�������</font></a></font></td>
</tr>
<tr> 
<td width="30%"><font color="<%=Forum_body(7)%>"><b>����</B>����</font></td>
<td width="70%"> 
<input type="text" name="Groupname" size="35" value="<%=rs("Groupname")%>">&nbsp;<input type="submit" name="Submit" value="�� ��">
<input type=hidden name=id value="<%=request("id")%>">
</td>
</tr>
<%set rs=nothing%>
<%else%>
<form method="POST" action=admin_group.asp?action=save>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>�������</b></font></td>
</tr>
<tr> 
<td width="30%"><font color="<%=Forum_body(7)%>"><b>����</B>����</font></td>
<td width="70%"> 
<input type="text" name="Groupname" size="35">&nbsp;<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<%end if%>
<tr bgcolor=<%=Forum_body(3)%>> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>"><b>��������</b></font></td>
</tr>
<%
	set rs=conn.execute("select * from GroupName")
	do while not rs.eof
%>
<tr> 
<td height="23" colspan="2" ><font color="<%=Forum_body(7)%>">
<a href="admin_group.asp?id=<%=rs("id")%>&action=edit"><font color="<%=Forum_body(7)%>">�޸�</font></a> | <a href="admin_group.asp?id=<%=rs("id")%>&action=del"><font color="<%=Forum_body(7)%>">ɾ��</font></a> | <%=rs("GroupName")%></font>
</td>
</tr>
<%
	rs.movenext
	loop
	set rs=nothing
%>
</table>
<%
end sub

sub savegroup()
	conn.execute("insert into GroupName (GroupName) values ('"&request("GroupName")&"')")
%>
<center><p><b>��ӳɹ���</b>
<%
end sub
sub savedit()
	conn.execute("update GroupName set GroupName='"&request("GroupName")&"' where id="&request("id"))
%>
<center><p><b>�޸ĳɹ���</b>
<%
end sub
sub del()
	conn.execute("delete from GroupName where id="&request("id"))
%>
<center><p><b>ɾ���ɹ���</b>
<%
end sub
%>