<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	dim admin_flag
	admin_flag="13"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if Request("action") = "unite" then
		call unite()
		else
		call boardinfo()
		end if
		conn.close
		set conn=nothing
	end if

sub boardinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height=25>�ϲ���̳����
	</th>
	</tr>
	<form action=admin_boardunite.asp?action=unite method=post>
	<tr>
	<td class=forumrow>
	<B>�ϲ���̳ѡ��</B>��<BR>
<B>������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������</B><BR><BR>
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select boardid,boardtype,depth from board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " ����̳ "
		response.write "<select name=oldboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	sql="select boardid,boardtype,depth from board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " �ϲ��� "
		response.write "<select name=newboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write " <BR><BR><input type=submit name=Submit value=�ϲ���̳><BR><BR>"
%>
	</td>
	</tr>
	<tr>
	<td class=forumrow><B>ע������</B>��<BR><FONT COLOR="red">���в��������棬�����ز���</FONT><BR> ������ͬһ�������ڽ��в��������ܽ�һ������ϲ�����������̳�С�<BR>�ϲ�������ָ������̳�����߰�����������̳������ɾ�����������ӽ�ת�Ƶ�����ָ����Ŀ����̳��
	</td>
	</tr></form>
	</table>
<%
end sub

sub unite()
dim newboard
dim oldboard
dim ParentStr,iParentStr
dim depth,iParentID,child
if cint(request("newboard"))=cint(request("oldboard")) then
	response.write "�벻Ҫ����ͬ�����ڽ��в�����"
	exit sub
end if
newboard=clng(request("newboard"))
oldboard=clng(request("oldboard"))
'������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������
'�õ���ǰ����������̳
set rs=conn.execute("select ParentStr,Boardid,depth,ParentID,child from board where boardid="&oldboard)
if rs(0)="0" then
ParentStr=rs(1)
iParentID=rs(1)
else
ParentStr=rs(1) & "," & rs(0)
iParentID=rs(3)
end if
iParentStr=rs(1)
depth=rs(2)
child=rs(4)+1
i=0
set rs=conn.execute("select Boardid from board where boardid="&newboard&" and (ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%')")
if not (rs.eof and rs.bof) then
	response.write "���ܽ�һ������ϲ�����������̳��"
	exit sub
end if
'�õ���ǰ����������̳ID
set rs=conn.execute("select Boardid from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
if not (rs.eof and rs.bof) then
do while not rs.eof
	iParentStr=iParentStr & "," & rs(0)
	i=i+1
rs.movenext
loop
end if
if i>0 then
	ParentStr=iParentStr
else
	ParentStr=oldboard
end if
'������ԭ��������̳������
if depth>0 then
conn.execute("update board set child=child-"&child&" where boardid="&iparentid)
'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
for i=1 to depth
	'�õ��丸��ĸ���İ���ID
	set rs=conn.execute("select parentid from board where boardid="&iparentid)
	if not (rs.eof and rs.bof) then
		iparentid=rs(0)
		conn.execute("update board set child=child-"&child&" where boardid="&iparentid)
	end if
next
end if
'������̳��������
For i=0 to ubound(AllPostTable)
conn.execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid in ("&ParentStr&")")
Next
conn.execute("update topic set boardid="&newboard&" where boardid in ("&ParentStr&")")
conn.execute("update besttopic set boardid="&newboard&" where boardid in ("&ParentStr&")")
'ɾ�����ϲ���̳
set rs=conn.execute("select sum(lastbbsnum),sum(lasttopicnum),sum(todayNum) from board where boardid in ("&ParentStr&")")
conn.execute("delete from board where boardid in ("&ParentStr&")")
'��������̳���Ӽ���
dim trs
set trs=conn.execute("select ParentStr,boardid from board where boardid="&newboard)
if trs(0)="0" then
ParentStr=trs(1)
else
ParentStr=trs(0)
end if
conn.execute("update board set lastbbsnum=lastbbsnum+"&rs(0)&",lasttopicnum=lasttopicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid in ("&ParentStr&")")
response.write "�ϲ��ɹ����Ѿ������ϲ���̳��������ת�������ϲ���̳����������������̳����ҳ�������̳���ݡ�"
set rs=nothing
set trs=nothing
call cache_board()
end sub
sub cache_board()
'cache��������
myCache.name="BoardJumpList"
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
if boardid=rs(0) then
BoardJumpList = BoardJumpList & "selected"
end if
BoardJumpList = BoardJumpList & ">"
select case rs(2)
case 0
BoardJumpList = BoardJumpList & "��"
case 1
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
end select
if rs(2)>1 then
for i=2 to rs(2)
	BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
next
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;��"
end if
BoardJumpList = BoardJumpList & rs(1)&"</option>"
rs.MoveNext
loop
myCache.add BoardJumpList,dateadd("n",60,now)
set rs=nothing
'end cache
end sub
%>