<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="01"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
		call saveconst()
		elseif request("action")="restore" then
		call restore()
		else
		call consted()
		end if
		conn.close
		set conn=nothing
	end if


sub consted()
dim sel
%>
<form method="POST" action="admin_setting.asp?action=save">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height=25 colspan=2 align=left><B>˵��</B>����������ô�����������̳�����е����ã����������úͷ���̳�����������������ü��ײ�ͬ�ķ����Թ�������ͬ�ķ���̳ʹ�á�</td>
</tr>
<tr> 
<td height=25 colspan=2 align=left>
<table width=85% align=center>
<tr>
<td width=50% ><a name=top></a><a href="#setting1">[��̳����]</a></td>
<td width=50% ><a href="#setting17">[��ˢ�»���]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting3">[������Ϣ]</a></td>
<td width=50% ><a href="#setting4">[��ϵ����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting16">[����ѡ��]</a></td>
<td width=50% ><a href="#setting6">[���Ļ�ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting7">[��̳��ҳѡ��]</a></td>
<td width=50% ><a href="#setting8">[�û���ע��ѡ��]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting19">[��������]</a></td>
<td width=50% ><a href="#setting10">[ϵͳ����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting18">[��̳��ҳ����]</a></td>
<td width=50% ><a href="#setting12">[���ߺ��û���Դ]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting13">[�ʼ�ѡ��]</a></td>
<td width=50% ><a href="#setting14">[�ϴ�����]</a></td>
</tr>
<tr>
<td width=50% ><a href="#setting15">[�û�ѡ��(ǩ����ͷ�Ρ����е�)]</a></td>
<td width=50% ></td>
</tr>
</table>
</td>
</tr> 
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2  class="tableHeaderText" height=25>��ǰʹ�����ã��ɽ���ǰ���ñ��浽��������ģ���У�ѡ�м��ɣ�
</th></tr>
<tr> 
<td width="100%" colspan=2 class="forumrow" height=25><input type="checkbox" name="checkall" checked value="1">&nbsp;ȫѡ�����Ҫʹ�������ö�����ģ����Ч����ȫѡ������ȫѡ
</td>
</tr>
<tr> 
<td width="100%" colspan=2 class="forumrow">
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
	Forum_Setting=split(rs("Forum_setting"),",")
	if Forum_Setting(12)="" then Forum_Setting(12)="0"
	if Forum_Setting(13)="" then Forum_Setting(13)="0"
	if Forum_Setting(16)="" then Forum_Setting(16)="0"
	if Forum_Setting(17)="" then Forum_Setting(17)="1"
	if Forum_Setting(18)="" then Forum_Setting(18)="1"
	if Forum_Setting(27)="" then Forum_Setting(27)="0"
	if Forum_Setting(39)="" then Forum_Setting(39)=Forum_Setting(38)
	if Forum_Setting(43)="" then Forum_Setting(43)="0"
	if Forum_Setting(45)="" then Forum_Setting(45)="0"
	if Forum_Setting(58)="" then Forum_Setting(58)=Forum_Setting(57)
	if Forum_Setting(71)="" then Forum_Setting(71)="0"
	if Forum_Setting(72)="" then Forum_Setting(72)="00"
	uForum_Setting=ubound(Forum_Setting)
	if uForum_Setting<77 then
		redim preserve Forum_Setting(77)
		Forum_Setting(73)="0"
		Forum_Setting(74)="1"
		Forum_Setting(75)="1"
		Forum_Setting(76)="1"
		Forum_Setting(77)="3"
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
	Forum_Setting=split(rs("Forum_setting"),",")
	if Forum_Setting(12)="" then Forum_Setting(12)="0"
	if Forum_Setting(13)="" then Forum_Setting(13)="0"
	if Forum_Setting(16)="" then Forum_Setting(16)="0"
	if Forum_Setting(17)="" then Forum_Setting(17)="1"
	if Forum_Setting(18)="" then Forum_Setting(18)="1"
	if Forum_Setting(27)="" then Forum_Setting(27)="0"
	if Forum_Setting(39)="" then Forum_Setting(39)=Forum_Setting(38)
	if Forum_Setting(43)="" then Forum_Setting(43)="0"
	if Forum_Setting(45)="" then Forum_Setting(45)="0"
	if Forum_Setting(58)="" then Forum_Setting(58)="32"
	if Forum_Setting(71)="" then Forum_Setting(71)="0"
	if Forum_Setting(72)="" then Forum_Setting(72)="00"
	uForum_Setting=ubound(Forum_Setting)
	if uForum_Setting<77 then
		redim preserve Forum_Setting(77)
		Forum_Setting(73)="0"
		Forum_Setting(74)="1"
		Forum_Setting(75)="1"
		Forum_Setting(76)="1"
		Forum_Setting(77)="3"
	end if
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_setting.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
</td></tr>
<tr> 
<td width="100%" colspan=2 class="Forumrow"><B>��ǰʹ�ø����õķ���̳</B> | <a href="?action=restore&skinid=<%=skinid%>"><B>��ԭ��̳��������</B></a> <BR><BR>���·���̳��ʹ�ñ����ã������޸���̳ʹ�����ã��ɵ���̳��������������<BR>
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
<tr> 
<th height="25" colspan="2" align=left id=TableTitleLink><a name="setting1"></a><b>��̳����</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��������</U></td>
<td width="50%" class=Forumrow> 
<select size=1 name="forum_setting(9)">
<option value="lang_gb.asp">��������
</select>
</td>
</tr>
<!--shinzeal����cookies����ƭ����-->
<tr> 
<td width="50%" class=Forumrow><U>Cookies��ƭ����</U><BR>�������Ҫ��ֹCookies��ƭ���뿪�����ǹ̶�IP�û������޷����ڱ���Cookies��Ϣ��</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(72)" value=0 <%if cint(Forum_Setting(72))=0 then%>checked<%end if%>>�ر�
<input type=radio name="Forum_Setting(72)" value=1 <%if cint(Forum_Setting(72))=1 then%>checked<%end if%>>�԰������Ͽ���
<input type=radio name="Forum_Setting(72)" value=2 <%if cint(Forum_Setting(72))=2 then%>checked<%end if%>>�������û�����
</td>
</tr>
<!--shinzeal������-->

<tr> 
<td width="50%" class=Forumrow><U>��̳��ǰ״̬</U><BR>ά���ڼ�����ùر���̳ʹ��</td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(21)" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(21)" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>�ر�&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>ά��˵��</U><BR>����̳�ر��������ʾ��֧��html�﷨</td>
<td width="50%" class=Forumrow> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>�Ƿ����ö�ʱ������̳</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="forum_setting(69)" value=0 <%if cint(Forum_Setting(69))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(69)" value=1 <%if cint(Forum_Setting(69))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳����ʱ��</U><BR>��ȷ�����Ѿ��������ö�ʱ���ع���<BR>��ֹСʱ�÷��š�|���ֿ�</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="forum_setting(70)" value="<%=forum_setting(70)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳģʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(4)" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>�°���ʽ&nbsp;
<input type=radio name="Forum_Setting(4)" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>�ɰ���ʽ&nbsp;
</td>
</tr>
<td width="50%" class=Forumrow> <U>�����˵�ģʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(71)" value=0 <%if cint(Forum_Setting(71))=0 then%>checked<%end if%>>�����̶�&nbsp;
<input type=radio name="Forum_Setting(71)" value=1 <%if cint(Forum_Setting(71))=1 then%>checked<%end if%>>��໬��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>��̳������Ϣ</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��url</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳�Ľ�վ����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CreateDate" size="35" value="<%=Forum_info(12)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳURL</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��ҳLogo��ַ</U><BR>��ʾ����̳��ҳ�������̳���û����дlogo��ַ����ʹ�ø�����</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳ͼƬĿ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��������Ŀ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��������Ŀ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="emoturl" size="35" value="<%=Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳ͷ��Ŀ¼</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="UserFaceurl" size="35" value="<%=Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��Ȩ��Ϣ</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting4"></a><b>��̳��ϵ����</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����ԱEmail</U><BR>���û������ʼ�ʱ����ʾ����ԴEmail��Ϣ</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>���Ļ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�¶���Ϣ��������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(10)" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(10)" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting7"></a><b>��̳��ҳѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��ҳ��ʾ��̳���</U><BR>0����һ����1����2�����Դ�����<BR>���ù������̳��Ƚ�Ӱ����̳�������ܣ�������Լ���̳��������ã���������Ϊ1</td>
<td width="50%" class=Forumrow> 
<input type=text size=10 name="forum_setting(5)" value="<%=forum_setting(5)%>"> ��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ٵ�¼λ��</U></td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(31)">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>����
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>�ײ�
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>����ʾ
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���ʾ�����ջ�Ա</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(29)" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(29)" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���ʾ��������Ա</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(43)" value=0 <%if cint(Forum_Setting(43))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(43)" value=1 <%if cint(Forum_Setting(43))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ�Ƿ���ʾ����б�(ͣ��)</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(12)" value=0 <%if cint(Forum_Setting(12))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(12)" value=1 <%if cint(Forum_Setting(12))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ�Ƿ���ʾ����վ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(13)" value=0 <%if cint(Forum_Setting(13))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(13)" value=1 <%if cint(Forum_Setting(13))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ�Ƿ���ʾ����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(16)" value=0 <%if cint(Forum_Setting(16))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%if cint(Forum_Setting(16))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ǵ�һ�е���ʾ��ʽ</U></td>
<td width="50%" class=Forumrow><select name="Forum_Setting(17)"> 
<option value=1 <%if cint(Forum_Setting(17))=1 then%>selected<%end if%>>���շ�����</option>
<option value=2 <%if cint(Forum_Setting(17))=2 then%>selected<%end if%>>���ܷ�����</option>
<option value=3 <%if cint(Forum_Setting(17))=3 then%>selected<%end if%>>���·�����</option>
<option value=4 <%if cint(Forum_Setting(17))=4 then%>selected<%end if%>>���귢����</option>
<option value=5 <%if cint(Forum_Setting(17))=5 then%>selected<%end if%>>��������</option>
<option value=6 <%if cint(Forum_Setting(17))=6 then%>selected<%end if%>>���������</option>
<option value=7 <%if cint(Forum_Setting(17))=7 then%>selected<%end if%>>���Ů����</option>
<option value=8 <%if cint(Forum_Setting(17))=8 then%>selected<%end if%>>���յ÷���</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ǵڶ��е���ʾ��ʽ</U></td>
<td width="50%" class=Forumrow><select name="Forum_Setting(18)"> 
<option value=1 <%if cint(Forum_Setting(18))=1 then%>selected<%end if%>>���շ�����</option>
<option value=2 <%if cint(Forum_Setting(18))=2 then%>selected<%end if%>>���ܷ�����</option>
<option value=3 <%if cint(Forum_Setting(18))=3 then%>selected<%end if%>>���·�����</option>
<option value=4 <%if cint(Forum_Setting(18))=4 then%>selected<%end if%>>���귢����</option>
<option value=5 <%if cint(Forum_Setting(18))=5 then%>selected<%end if%>>��������</option>
<option value=6 <%if cint(Forum_Setting(18))=6 then%>selected<%end if%>>���������</option>
<option value=7 <%if cint(Forum_Setting(18))=7 then%>selected<%end if%>>���Ů����</option>
<option value=8 <%if cint(Forum_Setting(18))=8 then%>selected<%end if%>>���յ÷���</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳ�Ƿ���ʾ������ϲ������̳��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(75)" value=0 <%if cint(Forum_Setting(75))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(75)" value=1 <%if cint(Forum_Setting(75))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ͼƬ�Ƿ��뵭��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(76)" value=0 <%if cint(Forum_Setting(76))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(76)" value=1 <%if cint(Forum_Setting(76))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ҳLogo������̳��ʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(77)" value=0 <%if cint(Forum_Setting(77))=0 then%>checked<%end if%>>������������&nbsp;
<input type=radio name="Forum_Setting(77)" value=1 <%if cint(Forum_Setting(77))=1 then%>checked<%end if%>>ֻ����������&nbsp;
<input type=radio name="Forum_Setting(77)" value=2 <%if cint(Forum_Setting(77))=2 then%>checked<%end if%>>ֻ����������&nbsp;
<input type=radio name="Forum_Setting(77)" value=3 <%if cint(Forum_Setting(77))=3 then%>checked<%end if%>>�ȹ����ֵ���&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>�û���ע��ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ��������û�ע��</U><BR>�رպ���̳������ע��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(37)" value=0 <%if cint(forum_setting(37))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if cint(forum_setting(37))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(40)" size="3" value="<%=forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(41)" size="3" value="<%=forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ͬһIPע����ʱ��</U><BR>�粻�����ƿ���д0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(22)" size="3" value="<%=Forum_Setting(22)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email֪ͨ����</U><BR>ȷ������վ��֧�ַ���mail������������Ϊϵͳ�������</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(23)" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(23)" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>һ��Emailֻ��ע��һ���ʺ�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(24)" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(24)" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ע����Ҫ����Ա��֤</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(25)" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(25)" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����ע����Ϣ�ʼ�</U><BR>��ȷ���������ʼ�����</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(47)" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(47)" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������Ż�ӭ��ע���û�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(46)" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(46)" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>ϵͳ����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ʱ��</U></td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ű���ʱʱ��</U><BR>Ĭ��Ϊ300��һ�㲻������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(1)" size="3" value="<%=Forum_Setting(1)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ���ʾҳ��ִ��ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(30)" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(30)" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>��ֹ���ʼ���ַ</U><BR>������ָ�����ʼ���ַ������ֹע�ᣬÿ���ʼ���ַ�á�|�����ŷָ�<BR>������֧��ģ����������������eway��ֹ������ֹeway@aspsky.net����eway@dvbbs.net����������ע��</td>
<td width="50%" class=Forumrow> 
<input type="text" name="Forum_Setting(52)" size="50" value="<%=Forum_Setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>��ˢ�¹�����Ч��ҳ��</U><BR>��ȷ�������˷�ˢ�¹���<BR>��ָ����ҳ�潫�з�ˢ�����ã��û����޶���ʱ���ڲ����ظ��򿪸�ҳ�棬����һ��������Դ���ĵ�����<BR>ÿ��ҳ�������á�|�����Ÿ���</td>
<td width="50%" class=Forumrow> 
<input type="text" name="Forum_Setting(64)" size="50" value="<%=Forum_Setting(64)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>���ߺ��û���Դ</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ʾ�û�IP</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(28)" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="Forum_Setting(28)" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ʾ�û���Դ</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�<BR>���������ܽ�������Դ</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(36)" value=0 <%if cint(forum_setting(36))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if cint(forum_setting(36))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������б���ʾ�û���ǰλ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(33)" value=0 <%if cint(forum_setting(33))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if cint(forum_setting(33))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������б���ʾ�û���¼�ͻʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(34)" value=0 <%if cint(forum_setting(34))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if cint(forum_setting(34))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������б���ʾ�û�������Ͳ���ϵͳ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(35)" value=0 <%if cint(forum_setting(35))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if cint(forum_setting(35))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ʾ��������</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(15)" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(15)" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ʾ�û�����</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(14)" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(14)" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ɾ������û�ʱ��</U><BR>������ɾ�����ٷ����ڲ���û�<BR>��λ�����ӣ�����������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(8)" size="3" value="<%=Forum_Setting(8)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����̳����ͬʱ������</U><BR>�粻�����ƣ�������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(26)" size="6" value="<%=Forum_Setting(26)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting13"></a><b>�ʼ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����ʼ����</U><BR>������ķ�������֧�������������ѡ��֧��</td>
<td width="50%" class=Forumrow>  
<select name="Forum_Setting(2)">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>��֧�� 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>SMTP Server��ַ</U><BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
<td width="50%" class=Forumrow>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting14"></a><b>�ϴ�����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ͷ���ϴ�</U></td>
<td width="50%" class=Forumrow> 
<input type=radio name="Forum_Setting(7)" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(7)" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>��������ͷ���ļ���С</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(56)" size="6" value="<%=forum_setting(56)%>">&nbsp;K
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>�û�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������ǩ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(42)" value=0 <%if forum_setting(42)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if forum_setting(42)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����û�ʹ��ͷ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(53)" value=0 <%if forum_setting(53)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if forum_setting(53)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>���ͷ����</U><BR>��������Ϊͷ��������</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(57)" size="6" value="<%=forum_setting(57)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>���ͷ��߶�</U><BR>��������Ϊͷ������߶�</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(58)" size="6" value="<%=forum_setting(58)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Ĭ��ͷ����</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(38)" size="6" value="<%=forum_setting(38)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Ĭ��ͷ��߶�</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(39)" size="6" value="<%=forum_setting(39)%>">&nbsp;����
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>ʹ���Զ���ͷ������ٷ�����</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=forum_setting(54)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������վ���ϴ�ͷ��</U><BR>�����Ƿ����ֱ��ʹ��http..������url��ֱ����ʾͷ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(55)" value=0 <%if forum_setting(55)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if forum_setting(55)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ǩ���Ƿ���UBB����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(65)" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(65)" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ǩ���Ƿ���HTML����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(66)" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(66)" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û��Ƿ�����ͼ��ǩ</U><BR>����ͼƬ��flash����ý���</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(67)" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(67)" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>�û������б����</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="Forum_Setting(68)" size="6" value="<%=Forum_Setting(68)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ͷ��</U><BR>�Ƿ������û��Զ���ͷ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(6)" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(6)" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ͷ����󳤶�</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(59)" size="6" value="<%=forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ�����ٷ�����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(60)" size="6" value="<%=forum_setting(60)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ��ע����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(61)" size="6" value="<%=forum_setting(61)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ������������������һ������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(62)" value=0 <%if cint(Forum_Setting(62))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(62)" value=1 <%if cint(Forum_Setting(62))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ����Ҫ���εĴ���</U><BR>ÿ�������ַ��á�|�����Ÿ���</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(63)" size="50" value="<%=forum_setting(63)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>��ˢ�»���</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ˢ�»���</U><BR>��ѡ�������д���������ˢ��ʱ��<BR>�԰����͹���Ա��Ч</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(19)" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(19)" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ˢ��ʱ����</U><BR>��д����Ŀ��ȷ�������˷�ˢ�»���<BR>���������б����ʾ����ҳ��������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(20)" size="3" value="<%=Forum_Setting(20)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>��̳��ҳ����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ÿҳ��ʾ����¼</U><BR>������̳���кͷ�ҳ�йص���Ŀ�������б��������ӳ��⣩</td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_Setting(11)" size="3" value="<%=Forum_Setting(11)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>����ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��Ϊ���Ż�����������ֵ</U><BR>��׼Ϊ����ظ���</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(44)" size="3" value="<%=forum_setting(44)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�༭����������ʾ"��xxx��yyy�༭"����Ϣ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(48)" value=0 <%if cint(forum_setting(48))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if cint(forum_setting(48))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����Ա�༭����ʾ"��XXX�༭"����Ϣ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(49)" value=0 <%if cint(forum_setting(49))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if cint(forum_setting(49))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ȴ�"��XXX�༭"��Ϣ��ʾ��ʱ��</U><BR>�����û��༭�Լ������Ӷ��������ӵײ���ʾ"��XXX�༭"��Ϣ��ʱ��(�Է���Ϊ��λ)</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(50)" size="3" value="<%=forum_setting(50)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�༭����ʱ��</U><BR>�༭�������ӵ�ʱ������(�Է���Ϊ��λ, 1����1440����) �������ʱ������, ֻ�й���Ա�Ͱ������ܱ༭��ɾ������. �������ʹ�������, ������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(51)" size="3" value="<%=forum_setting(51)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ����ü�¼�����Ķ��û�����</U><BR>ֻ���û�����ʱѡ���¼�Ķ������û�����Ч<BR>���������ܽ����Ĳ���ϵͳ��Դ<BR>�����û�������Ա�Ͱ����ɿ��������Ķ��û��б�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(3)" value=0 <%if cint(forum_setting(3))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(3)" value=1 <%if cint(forum_setting(3))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����й��ģʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(27)" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>��̬���&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>��������&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����л�Աͷ����ʾ��ʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(72_1)" value=0 <%if cint(left(Forum_Setting(72),1))=0 then%>checked<%end if%>>������&nbsp;
<input type=radio name="Forum_Setting(72_1)" value=1 <%if cint(left(Forum_Setting(72),1))=1 then%>checked<%end if%>>С����&nbsp;
<input type=radio name="Forum_Setting(72_1)" value=2 <%if cint(left(Forum_Setting(72),1))=2 then%>checked<%end if%>>ͷ��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�����л�Ա������ʾ��ʽ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(72_2)" value=0 <%if cint(right(Forum_Setting(72),1))=0 then%>checked<%end if%>>��ϸ&nbsp;
<input type=radio name="Forum_Setting(72_2)" value=1 <%if cint(right(Forum_Setting(72),1))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������Ƿ���ʾ���͹�Ʊ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(45)" value=0 <%if cint(Forum_Setting(45))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(45)" value=1 <%if cint(Forum_Setting(45))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ�����������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(73)" value=0 <%if cint(Forum_Setting(73))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="Forum_Setting(73)" value=1 <%if cint(Forum_Setting(73))=1 then%>checked<%end if%>>����������&nbsp;
<input type=radio name="Forum_Setting(73)" value=2 <%if cint(Forum_Setting(73))=2 then%>checked<%end if%>>���ƻظ���&nbsp;
<input type=radio name="Forum_Setting(73)" value=3 <%if cint(Forum_Setting(73))=3 then%>checked<%end if%>>����������&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������Ʊ���</U><BR>ÿ��¼��̳һ����Է�������Ӹ���</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(74)" size="3" value="<%=forum_setting(74)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>��������</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ�����̳����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_Setting(32)" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(32)" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>

<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
	dim Forum_copyright
	if trim(request("skinid"))="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
	else
		Forum_Setting=request.Form("Forum_Setting(0)") & "," & request.Form("Forum_Setting(1)") & "," & request.Form("Forum_Setting(2)") & "," & request.Form("Forum_Setting(3)") & "," & request.Form("Forum_Setting(4)") & "," & request.Form("Forum_Setting(5)") & "," & request.Form("Forum_Setting(6)") & "," & request.Form("Forum_Setting(7)") & "," & request.Form("Forum_Setting(8)") & "," & request.Form("Forum_Setting(9)") & "," & request.Form("Forum_Setting(10)") & "," & request.Form("Forum_Setting(11)") & "," & request.Form("Forum_Setting(12)") & "," & request.Form("Forum_Setting(13)") & "," & request.Form("Forum_Setting(14)") & "," & request.Form("Forum_Setting(15)") & "," & request.Form("Forum_Setting(16)") & "," & request.Form("Forum_Setting(17)") & "," & request.Form("Forum_Setting(18)") & "," & request.Form("Forum_Setting(19)") & "," & request.Form("Forum_Setting(20)") & "," & request.Form("Forum_Setting(21)") & "," & request.Form("Forum_Setting(22)") & "," & request.Form("Forum_Setting(23)") & "," & request.Form("Forum_Setting(24)") & "," & request.Form("Forum_Setting(25)") & "," & request.Form("Forum_Setting(26)") & "," & request.Form("Forum_Setting(27)") & "," & request.Form("Forum_Setting(28)") & "," & request.Form("Forum_Setting(29)") & "," & request.Form("Forum_Setting(30)") & "," & request.Form("Forum_Setting(31)") & "," & request.Form("Forum_Setting(32)") & "," & request.Form("Forum_Setting(33)") & "," & request.Form("Forum_Setting(34)") & "," & request.Form("Forum_Setting(35)") & "," & request.Form("Forum_Setting(36)") & "," & request.Form("Forum_Setting(37)") & "," & request.Form("Forum_Setting(38)") & "," & request.Form("Forum_Setting(39)") & "," & request.Form("Forum_Setting(40)") & "," & request.Form("Forum_Setting(41)") & "," & request.Form("Forum_Setting(42)") & "," & request.Form("Forum_Setting(43)") & "," & request.Form("Forum_Setting(44)") & "," & request.Form("Forum_Setting(45)") & "," & request.Form("Forum_Setting(46)") & "," & request.Form("Forum_Setting(47)") & "," & request.Form("Forum_Setting(48)") & "," & request.Form("Forum_Setting(49)") & "," & request.Form("Forum_Setting(50)") & "," & request.Form("Forum_Setting(51)") & "," & request.Form("Forum_Setting(52)") & "," & request.Form("Forum_Setting(53)") & "," & request.Form("Forum_Setting(54)") & "," & request.Form("Forum_Setting(55)") & "," & request.Form("Forum_Setting(56)") & "," & request.Form("Forum_Setting(57)") & "," & request.Form("Forum_Setting(58)") & "," & request.Form("Forum_Setting(59)") & "," & request.Form("Forum_Setting(60)") & "," & request.Form("Forum_Setting(61)") & "," & request.Form("Forum_Setting(62)") & "," & request.Form("Forum_Setting(63)") & "," & request.Form("Forum_Setting(64)") & "," & request.Form("Forum_Setting(65)") & "," & request.Form("Forum_Setting(66)") & "," & request.Form("Forum_Setting(67)") & "," & request.Form("Forum_Setting(68)") & "," & request.Form("Forum_Setting(69)") & "," & request.Form("Forum_Setting(70)") & "," & request.Form("Forum_Setting(71)") & "," & request.Form("Forum_Setting(72_1)") & request.Form("Forum_Setting(72_2)") & "," & request.Form("Forum_Setting(73)") & "," & request.Form("Forum_Setting(74)") & "," & request.Form("Forum_Setting(75)") & "," & request.Form("Forum_Setting(76)") & "," & request.Form("Forum_Setting(77)")
		Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT") & "," & Request("emoturl") & "," & Request("userfaceurl") & "," & Request("CreateDate")
		Forum_copyright=request("copyright")
		if request("checkall")="1" then
			sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"'"
			conn.execute(sql)
		else
			if request("skinid")<>"" then
				sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id in ( "&request("skinid")&" )"
				conn.execute(sql)
			end if
		end if
		response.write "���óɹ���"
		call cache_link(request.Form("Forum_Setting(77)"))
	end if
end sub

'�ָ�Ĭ������
sub restore()
Forum_info="�����ȷ���̳,http://www.dvbbs.net/,�����ȷ�,http://www.aspsky.net/,61.145.114.64,club@aspsky.net,images/LOGO.GIF,pic/,face/,����ʱ��,emot/,userface/"
Forum_Setting="0,300,1,1,0,1,1,1,40,lang_gb.asp,0,20,,,0,0,,,,0,30,0,40,0,1,0,9999,,1,0,1,1,1,1,1,1,0,1,32,32,0,10,1,,10,,1,1,1,1,10,10,0,1,0,1,500,120,,9,15,4,0,300,0,1,0,1,20,0,0|24,,"
conn.execute("update config set Forum_Setting='"&Forum_Setting&"',Forum_info='"&Forum_info&"' where id="&request("skinid"))
response.write "��ԭ���óɹ���"
end sub
%>