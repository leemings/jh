<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	dim admin_flag
	admin_flag="51,52"
	if not master or instr(session("flag"),"51")=0 or instr(session("flag"),"52")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
	dim sel

if request("action") = "savebadword" then
call savebadword()
else

%>

<form action="admin_badword.asp?action=savebadword" method=post>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2 align=left>��ǰʹ����ģ�壨�ɽ����ñ��浽����ģ���У�</th>
</tr>
<tr> 
<td width="100%" colspan=2 class="forumrow" height=25><input type="checkbox" name="checkall" checked value="1">&nbsp;ȫѡ�����Ҫʹ�������ö�����ģ����Ч����ȫѡ������ȫѡ
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
	else
	sel=""
	end if
	else
	sel=""
	end if
else
	if rs("id")=cint(request("skinid")) then
	sel="checked"
	badwords=rs("badwords")
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_badword.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%></font>
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
<%if request("reaction")="badword" then%>
<tr>
<th colspan=2 align=left height=23>���ӹ����ַ�</th>
</tr>
<tr>
<td class=forumrow width="20%">˵����</td>
<td class=forumrow width="80%">���ӹ����ַ����������������а��������ַ������ݣ�������Ҫ���˵��ַ������룬����ж���ַ��������á�|���ָ��������磺�����|�ҿ�|fuck</td>
</tr>
<tr>
<td class=forumrow width="20%">����������ַ�</td>
<td class=forumrow width="80%"><input type="text" name="badwords" value="<%=badwords%>" size="80"></td>
</tr>
<%elseif request("reaction")="splitreg" then%>
<tr>
<th colspan=2 align=left height=23>ע������ַ�</th>
</tr>
<tr>
<td class=forumrow width="20%">˵����</td>
<td class=forumrow width="80%">ע������ַ����������û�ע����������ַ������ݣ�������Ҫ���˵��ַ������룬����ж���ַ��������á�,���ָ��������磺ɳ̲,quest,ľ��</td>
</tr>
<tr>
<td class=forumrow width="20%">����������ַ�</td>
<td class=forumrow width="80%"><input type="text" name="splitwords" value="<%=splitwords%>" size="80"></td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=forumrow width="20%"></td>
<td width="80%" class=forumrow><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table>
</form>
<%end if%>
<%
end sub

sub savebadword()
	if trim(request("skinid"))="" then
		Founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
	else
		if request("reaction")="badword" then
			if request("checkall")="1" then
				sql = "update config set badwords='"&request("badwords")&"'"
				conn.execute(sql)
			else
				sql = "update config set badwords='"&request("badwords")&"' where id in ( "&request("skinid")&" )"
				conn.execute(sql)
			end if
		elseif request("reaction")="splitreg" then
			if request("checkall")="1" then
				sql = "update config set splitwords='"&request("splitwords")&"'"
				conn.execute(sql)
			else
				sql = "update config set splitwords='"&request("splitwords")&"' where id in ( "&request("skinid")&" )"
				conn.execute(sql)
			end if
		end if
		response.write "���³ɹ���"
	end if
end sub
%>