<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="md5.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	if not master or instr(session("flag"),"11")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
	Errmsg="<br><li>����İ�����Ϣ"
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
<th height="25" colspan="2" align=left>��̳�߼����� �� <%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>����̳��Ϊ������̳��������</U><BR>����Ѿ���������ʾ����������ת�Ƶ������̳<BR>ѡ���˸�������л�Ա�������ڱ��淢��/�����Ȳ���</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(43)" value=0 <%if cint(Board_Setting(43))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(43)" value=1 <%if cint(Board_Setting(43))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳�б���ʾ������̳���</U><BR>������̳��������̳��ʱ����Ч</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(39)" value=0 <%if cint(Board_Setting(39))=0 then%>checked<%end if%>>�б�&nbsp;
<input type=radio name="Board_Setting(39)" value=1 <%if cint(Board_Setting(39))=1 then%>checked<%end if%>>���&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳�б�����һ�а�����</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(41)" value="<%=Board_Setting(41)%>"> ��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳���</U><BR>��Ч����������ʾ��������ҳ��</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(24)" value=0 <%if cint(Board_Setting(24))=0 then%>checked<%end if%>>Ĭ�Ϸ��&nbsp;
<input type=radio name="Board_Setting(24)" value=1 <%if cint(Board_Setting(24))=1 then%>checked<%end if%>>���������&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>���������Ԥ������</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(25)" value="<%=Board_Setting(25)%>"> byte
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ�������̳</U><BR>������ֻ̳�й���Ա�͸ð�������ɽ�<BR></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(0)" value=0 <%if cint(Board_Setting(0))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(0)" value=1 <%if cint(Board_Setting(0))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ�������̳</U><BR>������ֻ̳�й���Ա�͸ð�������ɼ��ͽ���<BR>����û������̳Ȩ�޹�����û�Ȩ�޹������������û��ɼ��ͽ���<BR>�����ƶ�һ����̳����Ч</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(1)" value=0 <%if cint(Board_Setting(1))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(1)" value=1 <%if cint(Board_Setting(1))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ���֤��̳</U><BR>������ֻ̳�й���Ա�͸ð�������ɼ��ͽ���<BR>�û��ɽ�����趨�ڰ��������<BR>ֻ����֤�û��ɽ���</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(2)" value=0 <%if cint(Board_Setting(2))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(2)" value=1 <%if cint(Board_Setting(2))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ�VIP��̳</U><BR>VIP��ֻ̳��VIP��Ա�͹���Ա�ɼ��ͽ���<BR>�û������������VIP��Ա��Ȼ���ٽ���</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(51)" value=0 <%if cint(Board_Setting(51))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(51)" value=1 <%if cint(Board_Setting(51))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ��Ա���̳</U><BR>�Ա���ֻ̳��ָ���Ա�Ļ�Ա�͹���Ա�ɼ��ͽ���</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(52)" value=0 <%if cint(Board_Setting(52))=0 then%>checked<%end if%>>��ͨ��̳&nbsp;
<input type=radio name="Board_Setting(52)" value=1 <%if cint(Board_Setting(52))=1 then%>checked<%end if%>>Ů����̳&nbsp;
<input type=radio name="Board_Setting(52)" value=2 <%if cint(Board_Setting(52))=2 then%>checked<%end if%>>������̳&nbsp;
</td>
</tr>
<%'=============== ͬѧ¼ ��ʼ ===================%>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ�ͬѧ¼��̳</U><BR>ͬѧ¼��ֻ̳�м���ð༶�ĳ�Ա�͹���Ա���Խ���</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(50)" value=0 <%if cint(Board_Setting(50))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(50)" value=1 <%if cint(Board_Setting(50))=1 then%>checked<%end if%>>��&nbsp;&nbsp��ʼ��&nbsp;<input type=text size=15 name="boardmaster" value="<%if not isnull(rs("txlpd")) and rs("txlpd")<>"" then%><%=split(rs("txlpd"),"$")(3)%><%end if%>"><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ഴʼ���������|�ָ�
<input type="hidden" name="oldboardmaster" value="<%if not isnull(rs("txlpd")) and rs("txlpd")<>"" then%><%=split(rs("txlpd"),"$")(3)%><%end if%>">
</td>
</tr>
<%'=============== ͬѧ¼ ���� ===================%>
<tr> 
<td width="50%" class=Forumrow>
<U>��������ƶ�</U><BR>����������Ա�Ϳ���Ȩ���û��ɽ����������<BR>����������Ա�Ϳ���Ȩ���û���ֱ�ӷ���<BR>һ���û�����˺����ӷ��ɼ�</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(3)" value=0 <%if cint(Board_Setting(3))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(3)" value=1 <%if cint(Board_Setting(3))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ���ð����̳��ƶ�</U><BR>������ø��ƶȣ����ϼ���̳�����ɹ����¼���̳�����Ϣ</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(40)" value=0 <%if cint(Board_Setting(40))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(40)" value=1 <%if cint(Board_Setting(40))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ���ʾ�¼���̳����</U><BR>������¼���̳����ʾ������ط���������</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(42)" value=0 <%if cint(Board_Setting(42))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(42)" value=1 <%if cint(Board_Setting(42))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>����ͬʱ������</U><BR>������������Ϊ0</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(18)" value="<%=Board_Setting(18)%>"> ��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ����ö�ʱ������̳</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(21)" value=0 <%if cint(Board_Setting(21))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(21)" value=1 <%if cint(Board_Setting(21))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳����ʱ��</U><BR>��ȷ�����Ѿ��������ö�ʱ���ع���<BR>��ֹСʱ�÷��š�|���ֿ�</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(22)" value="<%=Board_Setting(22)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����б�ÿҳ��¼��</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(26)" value="<%=Board_Setting(26)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�������ÿҳ��¼��</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(27)" value="<%=Board_Setting(27)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>������������ɾ������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(33)" value=0 <%if cint(Board_Setting(33))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(33)" value=1 <%if cint(Board_Setting(33))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����������޸�CSS����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(34)" value=0 <%if cint(Board_Setting(34))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(34)" value=1 <%if cint(Board_Setting(34))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>���а��������޸�CSS����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(35)" value=0 <%if cint(Board_Setting(35))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(35)" value=1 <%if cint(Board_Setting(35))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ����̳�¼��еĲ����߱���</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(36)" value=0 <%if cint(Board_Setting(36))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(36)" value=1 <%if cint(Board_Setting(36))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����б�Ĭ�϶�ȡ������</U></td>
<td width="50%" class=Forumrow> 
<select size="0" name="Board_Setting(37)">
<option value="1"<%if cint(Board_Setting(37))=1 then%> selected<%end if%>>ȫ����ʾ����</option>
<option value="2"<%if cint(Board_Setting(37))=2 then%> selected<%end if%>>����������</option>
<option value="3"<%if cint(Board_Setting(37))=3 then%> selected<%end if%>>����������</option>
<option value="4"<%if cint(Board_Setting(37))=4 then%> selected<%end if%>>һ��������</option>
<option value="5"<%if cint(Board_Setting(37))=5 then%> selected<%end if%>>����������</option>
<option value="6"<%if cint(Board_Setting(37))=6 then%> selected<%end if%>>����������</option>
<option value="7"<%if cint(Board_Setting(37))=7 then%> selected<%end if%>>����������</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����б�Ĭ������ʽ</U></td>
<td width="50%" class=Forumrow> 
<select size="1" name="Board_Setting(38)">
<option value="1"<%if cint(Board_Setting(38))=1 then%> selected<%end if%>>���ظ�ʱ��</option>
<option value="2"<%if cint(Board_Setting(38))=2 then%> selected<%end if%>>����ʱ��</option>
<option value="3"<%if cint(Board_Setting(38))=3 then%> selected<%end if%>>����</option>
<option value="4"<%if cint(Board_Setting(38))=4 then%> selected<%end if%>>������</option>
<option value="5"<%if cint(Board_Setting(38))=5 then%> selected<%end if%>>�ظ���</option>
<option value="6"<%if cint(Board_Setting(38))=6 then%> selected<%end if%>>�����</option>
</select>
</td>
</tr>
<tr> 
<th height="25" colspan="2" align=left>�����������</th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>����ģʽ</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(4)" value=0 <%if cint(Board_Setting(4))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(4)" value=1 <%if cint(Board_Setting(4))=1 then%>checked<%end if%>>�߼�&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>HTML�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(5)" value=0 <%if cint(Board_Setting(5))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(5)" value=1 <%if cint(Board_Setting(5))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>UBB�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(6)" value=0 <%if cint(Board_Setting(6))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(6)" value=1 <%if cint(Board_Setting(6))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��ͼ��ǩ</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(7)" value=0 <%if cint(Board_Setting(7))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(7)" value=1 <%if cint(Board_Setting(7))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����ǩ</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(8)" value=0 <%if cint(Board_Setting(8))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(8)" value=1 <%if cint(Board_Setting(8))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��ý���ǩ</U><BR>����Flash,RM,AVI��</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(9)" value=0 <%if cint(Board_Setting(9))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(9)" value=1 <%if cint(Board_Setting(9))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ž�Ǯ��</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(10)" value=0 <%if cint(Board_Setting(10))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(10)" value=1 <%if cint(Board_Setting(10))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ż�����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(11)" value=0 <%if cint(Board_Setting(11))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(11)" value=1 <%if cint(Board_Setting(11))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(12)" value=0 <%if cint(Board_Setting(12))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(12)" value=1 <%if cint(Board_Setting(12))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(13)" value=0 <%if cint(Board_Setting(13))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(13)" value=1 <%if cint(Board_Setting(13))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(14)" value=0 <%if cint(Board_Setting(14))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(14)" value=1 <%if cint(Board_Setting(14))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Żظ���</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(15)" value=0 <%if cint(Board_Setting(15))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(15)" value=1 <%if cint(Board_Setting(15))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷ų�����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(23)" value=0 <%if cint(Board_Setting(23))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(23)" value=1 <%if cint(Board_Setting(23))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(53)" value=0 <%if cint(Board_Setting(53))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(53)" value=1 <%if cint(Board_Setting(53))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ż�Ա��</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(54)" value=0 <%if cint(Board_Setting(54))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(54)" value=1 <%if cint(Board_Setting(54))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ŷ�����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(55)" value=0 <%if cint(Board_Setting(55))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(55)" value=1 <%if cint(Board_Setting(55))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷��Ա���</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(56)" value=0 <%if cint(Board_Setting(56))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(56)" value=1 <%if cint(Board_Setting(56))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ÿ߼���</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(57)" value=0 <%if cint(Board_Setting(57))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(57)" value=1 <%if cint(Board_Setting(57))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(58)" value=0 <%if cint(Board_Setting(58))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(58)" value=1 <%if cint(Board_Setting(58))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(59)" value=0 <%if cint(Board_Setting(59))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(59)" value=1 <%if cint(Board_Setting(59))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷���ʾ��</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(60)" value=0 <%if cint(Board_Setting(60))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(60)" value=1 <%if cint(Board_Setting(60))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(61)" value=0 <%if cint(Board_Setting(61))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(61)" value=1 <%if cint(Board_Setting(61))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷ŷ�����������</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(44)" value=0 <%if cint(Board_Setting(44))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(44)" value=1 <%if cint(Board_Setting(44))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��������Ĭ��״̬</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(45)" value=0 <%if cint(Board_Setting(45))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(45)" value=1 <%if cint(Board_Setting(45))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷ŷ�������Żظ�����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(46)" value=0 <%if cint(Board_Setting(46))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(46)" value=1 <%if cint(Board_Setting(46))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��������Żظ���Ĭ��״̬</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(47)" value=0 <%if cint(Board_Setting(47))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(47)" value=1 <%if cint(Board_Setting(47))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ŷ�������ע�⹦��</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(48)" value=0 <%if cint(Board_Setting(48))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(48)" value=1 <%if cint(Board_Setting(48))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ񿪷Ű������ֹ���</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(49)" value=0 <%if cint(Board_Setting(49))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(49)" value=1 <%if cint(Board_Setting(49))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�����󷵻�</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(17)" value=1 <%if cint(Board_Setting(17))=1 then%>checked<%end if%>>��ҳ&nbsp;
<input type=radio name="Board_Setting(17)" value=2 <%if cint(Board_Setting(17))=2 then%>checked<%end if%>>��̳&nbsp;
<input type=radio name="Board_Setting(17)" value=3 <%if cint(Board_Setting(17))=3 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>������������ֽ���</U><BR>1024�ֽڵ���1K</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(16)" value="<%=Board_Setting(16)%>"> �ֽ�
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>���������ֺ�</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(28)" value="<%=Board_Setting(28)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>���������м��</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(29)" value="<%=Board_Setting(29)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�ϴ��ļ�����</U><BR>ÿ���ļ������á�|���ŷֿ�</td>
<td width="50%" class=Forumrow> 
<input type=text size=50 name="Board_Setting(19)" value="<%=Board_Setting(19)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ����÷���ˮ����</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Board_Setting(30)" value=0 <%if cint(Board_Setting(30))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(30)" value=1 <%if cint(Board_Setting(30))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>ÿ�η������</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(31)" value="<%=Board_Setting(31)%>"> ��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>���ͶƱ��Ŀ</U></td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="Board_Setting(32)" value="<%=Board_Setting(32)%>"> ��
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type=hidden value="0" name="Board_Setting(20)">
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
	if not isnumeric(request("editid")) then
		Errmsg=Errmsg+"<br>"+"<li>����İ������"
		call dvbbs_error()
		exit sub
	else
		'============= ͬѧ¼ ��ʼ ==================
		dim txlpd,txlpd1
		if cint(request("Board_Setting(50)"))=1 and request("BoardMaster")="" then
			Errmsg=Errmsg+"<br>"+"<li>�����ͬѧ¼��ʼ�˲���"
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
		'============= ͬѧ¼ ���� ==================
		response.write "���óɹ���<a href=admin_boardsetting.asp?editid="&request("editid")&">���ذ���߼�����</a>"
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
			str=str&"������������û���<b>" &arr(i) &"</b> ���룺<b>"& pw &"</b><br><br>"
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

	'�ж�ԭ���������������Ƿ񻹵��ΰ�������û�е����򳷻����û�ְλ
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