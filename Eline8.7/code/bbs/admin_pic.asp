<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	dim admin_flag
	admin_flag="42"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
			call saveconst()
		else
			call consted()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub consted()
dim sel
%>
<form method="POST" action=admin_pic.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�<br>4������ͼƬ����������̳picĿ¼�У���Ҫ����Ҳ�뽫ͼ���ڸ�Ŀ¼
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 align=left>��ǰʹ����ģ�壨�ɽ����ñ��浽����ģ���У�</th>
</tr>
<tr>
<td colspan=3 class=forumrow>
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
	Forum_pic=split(rs("Forum_pic"),",")
	uForum_Pic=ubound(Forum_Pic)
	if uForum_Pic<16 then
		redim preserve Forum_Pic(16)
		Forum_Pic(16)="ao2.gif"
	end if
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_ubb=split(rs("Forum_ubb"),",")
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
	Forum_pic=split(rs("Forum_pic"),",")
	uForum_Pic=ubound(Forum_Pic)
	if uForum_Pic<16 then
		redim preserve Forum_Pic(16)
		Forum_Pic(16)="ao2.gif"
	end if
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_ubb=split(rs("Forum_ubb"),",")
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_pic.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=forumrow><B>��ǰʹ�ø�ģ��ķ���̳</B><BR>���·���̳��ʹ�ñ�ģ�����ã������޸���̳ʹ��ģ�壬�ɵ���̳��������������<BR>
<select size=1>
<%
set rs = server.CreateObject ("adodb.recordset")
sql="select * from board where sid="&skinid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<option>û����̳ʹ�ø�ģ��</option>"
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
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>����ͼƬ����</th>
</tr>
<tr> 
<td width="40%" class=forumrow>��̳����Ա</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(0)" size="20" value="<%=Forum_pic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(16)" size="20" value="<%=Forum_pic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��̳����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(1)" size="20" value="<%=Forum_pic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��̳���</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(2)" size="20" value="<%=Forum_pic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��ͨ��Ա</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(3)" size="20" value="<%=Forum_pic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���˻������Ա</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(4)" size="20" value="<%=Forum_pic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>ͻ����ʾ�Լ�����ɫ</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(5)" size="20" value="<%=Forum_pic(5)%>">&nbsp;&nbsp;<font color=<%=Forum_pic(5)%>>�Լ�</font>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>��̳ͼ��</th>
</tr>
<tr> 
<td width="40%" class=forumrow>������̳������������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(7)" size="20" value="<%=Forum_pic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>������̳������������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(6)" size="20" value="<%=Forum_pic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��֤��̳������������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(15)" size="20" value="<%=Forum_pic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��֤��̳������������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(14)" size="20" value="<%=Forum_pic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>������̳</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(0)" size="20" value="<%=Forum_boardpic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б����˵��ǰ���ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(8)" size="20" value="<%=Forum_pic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б��н��շ�����ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(9)" size="20" value="<%=Forum_pic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б����ܷ�����ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(10)" size="20" value="<%=Forum_pic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б�����������ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(11)" size="20" value="<%=Forum_pic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��̳����ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(12)" size="20" value="<%=Forum_pic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б��Ϸ������б�ͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_pic(13)" size="20" value="<%=Forum_pic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_pic(13)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>��̳����ͼ��</th>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(1)" size="20" value="<%=Forum_boardpic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����ͶƱ</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(2)" size="20" value="<%=Forum_boardpic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ظ�����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(4)" size="20" value="<%=Forum_boardpic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>С�ֱ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(3)" size="20" value="<%=Forum_boardpic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(5)" size="20" value="<%=Forum_boardpic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ʻ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(6)" size="20" value="<%=Forum_boardpic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(7)" size="20" value="<%=Forum_boardpic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>û���¶���</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(8)" size="20" value="<%=Forum_boardpic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���µĶ���Ϣ</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(9)" size="20" value="<%=Forum_boardpic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ҷ��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(10)" size="20" value="<%=Forum_boardpic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����������ģʽ</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(11)" size="20" value="<%=Forum_boardpic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>ƽ�����������ģʽ</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(12)" size="20" value="<%=Forum_boardpic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��һƪ����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(13)" size="20" value="<%=Forum_boardpic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��һƪ����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(14)" size="20" value="<%=Forum_boardpic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>ˢ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(7)" size="20" value="<%=Forum_statePic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>չ������̳�������б�</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(15)" size="20" value="<%=Forum_boardpic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ر�չ������̳�������б�</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_boardpic(16)" size="20" value="<%=Forum_boardpic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_boardpic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(5)" size="20" value="<%=Forum_statePic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б��з�ҳͼ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(6)" size="20" value="<%=Forum_statePic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>������ʾ���ļ���</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(8)" size="20" value="<%=Forum_statePic(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�򿪵�����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(0)" size="20" value="<%=Forum_statePic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���ŵ�����</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(1)" size="20" value="<%=Forum_statePic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(2)" size="20" value="<%=Forum_statePic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�̶�������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(3)" size="20" value="<%=Forum_statePic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�̶ܹ�������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(9)" size="20" value="<%=Forum_statePic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(4)" size="20" value="<%=Forum_statePic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>ͶƱ������</td>
<td width="60%" class=forumrow> 
<input type="text" name="Forum_statePic(12)" size="20" value="<%=Forum_statePic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_statePic(12)%>>
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<th height="23" colspan="2" align=left>��̳����ͼ��</th>
</tr>
<tr> 
<td width="40%" class=forumrow>���浱ҳΪ�ļ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_save" size="20" value="<%=Forum_TopicPic(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���汾��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_report" size="20" value="<%=Forum_TopicPic(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��ʾ�ɴ�ӡ�İ汾</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_print" size="20" value="<%=Forum_TopicPic(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ѱ���������̳�ղ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_bbsfav" size="20" value="<%=Forum_TopicPic(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ѱ�������ʵ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPost" size="20" value="<%=Forum_TopicPic(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���ͱ�ҳ�������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_EmailPage" size="20" value="<%=Forum_TopicPic(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ѱ�������IE�ղ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_iefav" size="20" value="<%=Forum_TopicPic(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���Ͷ��Ÿ�����</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_sms" size="20" value="<%=Forum_TopicPic(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��Ϊ����</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_friend" size="20" value="<%=Forum_TopicPic(21)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(21)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����û���Ϣ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_userinfo" size="20" value="<%=Forum_TopicPic(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����û���������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_search" size="20" value="<%=Forum_TopicPic(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û�����</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_Email" size="20" value="<%=Forum_TopicPic(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û�OICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_oicq" size="20" value="<%=Forum_TopicPic(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û�ICQ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_icq" size="20" value="<%=Forum_TopicPic(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û�MSN</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_msn" size="20" value="<%=Forum_TopicPic(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û���ҳ</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_homepage" size="20" value="<%=Forum_TopicPic(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_quote" size="20" value="<%=Forum_TopicPic(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�༭����</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_edit" size="20" value="<%=Forum_TopicPic(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���ٻظ�</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_freply" size="20" value="<%=Forum_TopicPic(22)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(22)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>ɾ������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_delete" size="20" value="<%=Forum_TopicPic(17)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(17)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_copy" size="20" value="<%=Forum_TopicPic(18)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(18)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>���뾫������</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_isbest" size="20" value="<%=Forum_TopicPic(19)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(19)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�û�IP</td>
<td width="60%" class=forumrow> 
<input type="text" name="P_ip" size="20" value="<%=Forum_TopicPic(20)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_TopicPic(20)%>>
</td>
</tr>
<tr > 
<th height="23" colspan="2" align=left>��̳UBB��ǩͼ��</th>
</tr>
<tr> 
<td width="40%" class=forumrow>����</td>
<td width="60%" class=forumrow> 
<input type="text" name="bold" size="20" value="<%=Forum_ubb(0)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(0)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>б��</td>
<td width="60%" class=forumrow> 
<input type="text" name="italicize" size="20" value="<%=Forum_ubb(1)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(1)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�»���</td>
<td width="60%" class=forumrow> 
<input type="text" name="underline" size="20" value="<%=Forum_ubb(2)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(2)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>����</td>
<td width="60%" class=forumrow> 
<input type="text" name="center" size="20" value="<%=Forum_ubb(3)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(3)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>URL����</td>
<td width="60%" class=forumrow> 
<input type="text" name="url1" size="20" value="<%=Forum_ubb(4)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(4)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>Email��ַ</td>
<td width="60%" class=forumrow> 
<input type="text" name="email1" size="20" value="<%=Forum_ubb(5)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(5)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��ͼ</td>
<td width="60%" class=forumrow> 
<input type="text" name="image" size="20" value="<%=Forum_ubb(6)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(6)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��Flash</td>
<td width="60%" class=forumrow> 
<input type="text" name="swf" size="20" value="<%=Forum_ubb(7)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(7)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��ShockWave</td>
<td width="60%" class=forumrow> 
<input type="text" name="Shockwave" size="20" value="<%=Forum_ubb(8)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(8)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��RMӰƬ</td>
<td width="60%" class=forumrow> 
<input type="text" name="rm" size="20" value="<%=Forum_ubb(9)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(9)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��Media PlayerӰƬ</td>
<td width="60%" class=forumrow> 
<input type="text" name="mp" size="20" value="<%=Forum_ubb(10)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(10)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��QuickTimeӰƬ</td>
<td width="60%" class=forumrow> 
<input type="text" name="qt" size="20" value="<%=Forum_ubb(11)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(11)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="quote1" size="20" value="<%=Forum_ubb(12)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(12)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>������</td>
<td width="60%" class=forumrow> 
<input type="text" name="fly" size="20" value="<%=Forum_ubb(13)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(13)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�ƶ���</td>
<td width="60%" class=forumrow> 
<input type="text" name="move" size="20" value="<%=Forum_ubb(14)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(14)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>������</td>
<td width="60%" class=forumrow> 
<input type="text" name="glow" size="20" value="<%=Forum_ubb(15)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(15)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��Ӱ��</td>
<td width="60%" class=forumrow> 
<input type="text" name="shadow" size="20" value="<%=Forum_ubb(16)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(16)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>�����б�</td>
<td width="60%" class=forumrow> 
<input type="text" name="smile" size="20" value="<%=Forum_ubb(17)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(17)%>>
</td>
</tr>
<tr> 
<td width="40%" class=forumrow>��������</td>
<td width="60%" class=forumrow> 
<input type="text" name="bsound" size="20" value="<%=Forum_ubb(18)%>">&nbsp;&nbsp;<img src=<%=Forum_info(7)%><%=Forum_ubb(18)%>>
</td>
</tr>
<tr> 
<td colspan=2 align=center class=forumrow><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_pic=request("Forum_pic(0)") & "," & request("Forum_pic(1)") & "," & request("Forum_pic(2)") & "," & request("Forum_pic(3)") & "," & request("Forum_pic(4)") & "," & request("Forum_pic(5)") & "," & request("Forum_pic(6)") & "," & request("Forum_pic(7)") & "," & request("Forum_pic(8)") & "," & request("Forum_pic(9)") & "," & request("Forum_pic(10)") & "," & request("Forum_pic(11)") & "," & request("Forum_pic(12)") & "," & request("Forum_pic(13)") & "," & request("Forum_pic(14)") & "," & request("Forum_pic(15)") & "," & request("Forum_pic(16)")

Forum_boardpic=request("Forum_boardpic(0)") & "," & request("Forum_boardpic(1)") & "," & request("Forum_boardpic(2)") & "," & request("Forum_boardpic(3)") & "," & request("Forum_boardpic(4)") & "," & request("Forum_boardpic(5)") & "," & request("Forum_boardpic(6)") & "," & request("Forum_boardpic(7)") & "," & request("Forum_boardpic(8)") & "," & request("Forum_boardpic(9)") & "," & request("Forum_boardpic(10)") & "," & request("Forum_boardpic(11)") & "," & request("Forum_boardpic(12)") & "," & request("Forum_boardpic(13)") & "," & request("Forum_boardpic(14)") & "," & request("Forum_boardpic(15)") & "," & request("Forum_boardpic(16)")

Forum_TopicPic=request("P_save") & "," & request("P_report") & "," & request("P_print") & "," & request("P_EmailPost") & "," & request("P_bbsfav") & "," & request("P_EmailPage") & "," & request("P_iefav") & "," & request("P_sms") & "," & request("P_search") & "," & request("P_userinfo") & "," & request("P_Email") & "," & request("P_oicq") & "," & request("P_icq") & "," & request("P_msn") & "," & request("P_homepage") & "," & request("P_quote") & "," & request("P_edit") & "," & request("P_delete") & "," & request("P_copy") & "," & request("P_isbest") & "," & request("P_ip") & "," & request("P_friend") & "," & request("P_freply")

Forum_statePic=request("Forum_statePic(0)") & "," & request("Forum_statePic(1)") & "," & request("Forum_statePic(2)") & "," & request("Forum_statePic(3)") & "," & request("Forum_statePic(4)") & "," & request("Forum_statePic(5)") & "," & request("Forum_statePic(6)") & "," & request("Forum_statePic(7)") & "," & request("Forum_statePic(8)") & "," & request("Forum_statePic(9)") & "," & request("Forum_statePic(10)") & "," & request("Forum_statePic(11)") & "," & request("Forum_statePic(12)")

Forum_ubb=request("bold") & "," & request("italicize") & "," & request("underline") & "," & request("center") & "," & request("url1") & "," & request("email1") & "," & request("image") & "," & request("swf") & "," & request("Shockwave") & "," & request("rm") & "," & request("mp") & "," & request("qt") & "," & request("quote1") & "," & request("fly") & "," & request("move") & "," & request("glow") & "," & request("shadow") & "," & request("smile") & "," & request("bsound")
'response.write Forum_Topicpic & "<br>" & Forum_boardpic & "<br>" & Forum_statePic
if request("skinid")<>"" then
sql = "update config set Forum_pic='"&Forum_pic&"',Forum_boardpic='"&Forum_boardpic&"',Forum_TopicPic='"&Forum_TopicPic&"',Forum_statePic='"&Forum_statePic&"',Forum_ubb='"&Forum_ubb&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
response.write "���óɹ���<a href="&Request.ServerVariables("HTTP_REFERER")&" >�������</a>"
end if
end sub
%>