<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=md5.asp-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim str
	dim admin_flag
	admin_flag="11,12"
	if not master or instr(session("flag"),"11")=0 or instr(session("flag"),"12")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2 height=25>��̳����
</th>
</tr>
<tr>
<td class="forumrow" colspan=2>
<p><B>ע��</B>��<BR>��ɾ����̳ͬʱ��ɾ������̳���������ӣ�ɾ������ͬʱɾ��������̳���������ӣ� ����ʱ��������д����Ϣ��<BR>�����ѡ��<B>��λ���а���</B>�������а��涼����Ϊһ����̳�����ࣩ����ʱ����Ҫ���¶Ը���������й����Ļ������ã�<B>��Ҫ����ʹ�øù���</B>�����������˴�������ö��޷���ԭ����֮��Ĺ�ϵ�������ʱ��ʹ��
</td>
</tr>
<tr>
<td class="forumrow">
<B>��̳����ѡ��</B></td>
<td class="forumrow"><a href="admin_board.asp">��̳������ҳ</a> | <a href="admin_board.asp?action=add">�½���̳����</a> | <a href="?action=orders">һ����������</a> | <a href="?action=boardorders">N����������</a> | <a href="?action=RestoreBoard" onclick="{if(confirm('��λ���а��潫�����а���ָ���Ϊһ������࣬��λ��Ҫ�����а������½��й����Ļ������ã������ز�����ȷ����λ��?')){return true;}return false;}">��λ���а���</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"
	call savenew()
case "savedit"
	call savedit()
case "del"
	call del()
case "orders"
	call orders()
case "updatorders"
	call updateorders()
case "boardorders"
	call boardorders()
case "updatboardorders"
	call updateboardorders()
case "addclass"
	call addclass()
case "saveclass"
	call saveclass()
case "del1"
	call del1()
case "mode"
	call mode()
case "savemod"
	call savemod()
case "permission"
	call boardpermission()
case "editpermission"
	call editpermission()
case "RestoreBoard"
	call RestoreBoard()
case else
	call boardinfo()
end select
end sub

sub boardinfo()
Dim reBoard_Setting
%>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>����
</th>
</tr>
<%
sql="select * from board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
reBoard_Setting=split(rs("Board_setting"),",")
%>
<tr> 
<td height="25" width=35%  class="forumrow">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<td width=65% align=right class="forumrow"><a href="admin_board.asp?action=add&editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>��Ӱ���</U></font></a> | <a href="admin_board.asp?action=edit&editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>��������</U></font></a> | <a href="admin_BoardSetting.asp?editid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>�߼�����</U></font></a>
<%if reBoard_Setting(2)=1 then%>
| <a href="admin_board.asp?action=mode&boardid=<%=rs("boardid")%>"><font color="<%=Forum_body(14)%>"><U>��֤�û�</U></font></a>
<%end if%>
| <a href="admin_update.asp?action=updat&submit=������̳����&boardid=<%=rs("boardid")%>" title="�������ظ������������ظ���"><font color="<%=Forum_body(14)%>"><U>��������</U></font></a> | <a href="admin_update.asp?action=delboard&boardid=<%=rs("boardid")%>" onclick="{if(confirm('��ս���������̳�����������ڻ���վ��ȷ�������?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>���</U></font></a> | <%if rs("child")=0 then%><a href="admin_board.asp?action=del&editid=<%=rs("boardid")%>" onclick="{if(confirm('ɾ������������̳���������ӣ�ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>ɾ��</U></a><%else%><a href="#" onclick="{if(confirm('����̳����������̳��������ɾ����������̳����ɾ������̳��')){return true;}return false;}"><font color="<%=Forum_body(14)%>"><U>ɾ��</U></a><%end if%>&nbsp;</td>
</tr>
<%
rs.movenext
loop
set rs=nothing
%>
</table><BR><BR>
<%
end sub

sub add()
dim rs_c
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from board order by rootid,orders"
rs_c.open sql,conn,1,1
	dim boardnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	rs.close
%>
<form action ="admin_board.asp?action=savenew" method=post>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2><B>�������̳</th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">��̳����</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>�������</U></td>
<td width="48%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if request("editid")<>"" and clng(request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
��
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>ʹ������ģ��</U><BR>���ģ���а�����̳��ɫ�����á���桢ͼƬ<BR>����Ϣ</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>�������ģ��"
	else
	do while not rs_c.EOF
%>
<option value="<%=rs_c("id")%>"><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumRow">&nbsp;</td>
<td width="48%" class="forumRow"> 
<input type="submit" name="Submit" value="�����̳">
</td>
</tr>
</table>
</form>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c
sql = "select * from board order by rootid,orders"
set rs_c=conn.execute(sql)
sql = "select * from board where boardid="&request("editid")
set rs=conn.execute(sql)
%>        
 <form action ="admin_board.asp?action=savedit" method=post>       
<input type="hidden" name=editid value="<%=Request("editid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2>�༭��̳��<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow">��̳����</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardtype" size="35"  value="<%=rs("boardtype")%>">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="48%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"><%=rs("readme")%></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>�������</U><BR>������̳����ָ��Ϊ��ǰ����<BR>������̳����ָ��Ϊ��ǰ�����������̳</td>
<td width="48%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>><%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
��
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>ʹ������ģ��</U><BR>���ģ���а�����̳��ɫ�����á���桢ͼƬ<BR>����Ϣ</td>
<td width="48%" class="forumrow"> 
<select name=sid>
<%
	sql = "select * from config order by active desc"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>�������ģ��"
	else
	do while not rs_c.EOF
%>
<option value=<%=rs_c("id")%> <% if cint(rs("sid")) = rs_c("id") then%> selected <%end if%>><%=rs_c("skinname")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="48%" class="forumrow"> 
<input type="text" name="boardmaster" size="35"  value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="52%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="48%" class="forumrow">
<input type="text" name="indexIMG" size="35" value="<%=rs("indexIMG")%>">
</td>
</tr>
<tr> 
<td width="52%" height=24 class="forumrow">&nbsp;</td>
<td width="48%" class="forumrow"> 
<input type="submit" name="Submit" value="�ύ�޸�">
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub
sub mode()
dim boarduser
%>
 <form action ="admin_board.asp?action=savemod" method=post>
<table width="95%" class="tableBorder" cellspacing="1" cellpadding="1" align="center">
<tr bgcolor=<%=Forum_body(3)%>> 
<th width="52%" height=22>˵����</th>
<th width="48%">������</th>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow><B>��̳����</B></td>
<td width="48%" class=forumrow> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql="select boardid,boardtype,boarduser from board where boardid="&request("boardid")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "�ð��沢�����ڻ��߸ð��治�Ǽ��ܰ��档"
else
response.write rs(1)
response.write "<input type=hidden value="&rs(0)&" name=boardid>"
boarduser=rs(2)
end if
rs.close
set rs=nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=forumrow><B>��֤�û�</B>��<br>
ֻ���趨Ϊ��֤��̳����̳��Ҫ��д�ܹ�����ð�����û���ÿ����һ���û���ȷ���û�������̳�д��ڣ�ÿ���û�����<B>�س�</B>�ֿ�</font>
</td>
<td width="48%" class=forumrow> 
<textarea cols=35 rows=6 name="vipuser">
<%
if not isnull(boarduser) or boarduser<>"" then
	response.write replace(boarduser,",",chr(10))
end if
%>
</textarea>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow>&nbsp;</td>
<td width="48%" class=forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end sub

'����༭��̳��֤�û���Ϣ
'��ڣ��û��б��ַ���
sub savemod()
dim boarduser
dim boarduser_1
dim userlen
dim updateinfo
response.write "<p>��̳���óɹ���<br><br>"
if trim(request("vipuser"))<>"" then
	boarduser=request("vipuser")
	boarduser=split(boarduser,chr(13)&chr(10))
	for i = 0 to ubound(boarduser)
	if not (boarduser(i)="" or boarduser(i)=" ") then
		boarduser_1=""&boarduser_1&""&boarduser(i)&","
	end if
	next
	userlen=len(boarduser_1)
	if boarduser_1<>"" then
		boarduser=left(boarduser_1,userlen-1)
		response.write "<p>����û���"&boarduser&"<br><br>"
		updateinfo=" boarduser='"&boarduser&"' "
	else
		response.write "<p><font color=red>��û�������֤�û�</font><br><br>"
	end if
end if
conn.execute("update board set "&updateinfo&" where boardid="&request("boardid"))
end sub

'���������̳��Ϣ
sub savenew()
if request("boardtype")="" then
	Errmsg=Errmsg+"<br>"+"<li>��������̳���ơ�"
	Founderr=true
end if
if request("class")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ѡ����̳���ࡣ"
	Founderr=true
end if
if request("readme")="" then
	Errmsg=Errmsg+"<br>"+"<li>��������̳˵����"
	Founderr=true
end if
if founderr=true then
	response.write Errmsg
	exit sub
end if
dim boardid
dim rootid
dim parentid
dim depth
dim orders
dim Fboardmaster
dim maxrootid
dim parentstr
if request("class")<>"0" then
set rs=conn.execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from board where boardid="&request("class"))
rootid=rs(0)
parentid=rs(1)
depth=rs(2)
orders=rs(3)
if depth+1>20 then
	response.write "����̳�������ֻ����20������"
	exit sub
end if
parentstr=rs(5)
else
set rs=conn.execute("select max(rootid) from board")
maxrootid=rs(0)+1
if isnull(MaxRootID) then MaxRootID=1
end if
sql="select boardid from board where boardid="&request("newboardid")
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	response.write "������ָ���ͱ����̳һ������š�"
	exit sub
else
	boardid=request("newboardid")
end if

set rs = server.CreateObject ("adodb.recordset")
sql = "select * from board"
rs.Open sql,conn,1,3
rs.AddNew
if request("class")<>"0" then
rs("depth")=depth+1
rs("rootid")=rootid
rs("orders") = Request.form("newboardid")
rs("parentid") = Request.Form("class")
if ParentStr="0" then
rs("ParentStr")=Request.Form("class")
else
rs("ParentStr")=ParentStr & "," & Request.Form("class")
end if
else
rs("depth")=0
rs("rootid")=maxrootid
rs("orders")=0
rs("parentid")=0
rs("child")=0
rs("parentstr")=0
end if
rs("boardid") = Request.form("newboardid")
rs("boardtype") = Request.Form("boardtype")
rs("readme") = Request.form("readme")
rs("lasttopicnum") = 0
rs("lastbbsnum") = 0
rs("lasttopicnum") = 0
rs("todaynum") = 0
rs("LastPost")="$0$"&Now()&"$$$$$"
rs("Board_Setting")="0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,0|24,1,0,300,20,10,9,12,1,10,10,0,0,0,0,1,5,0,1,4,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,1,1,1,1,1"
rs("sid")=request.form("sid")
if Request("boardmaster")<>"" then
	rs("boardmaster") = Request.form("boardmaster")
end if
if request.form("indexIMG")<>"" then
	rs("indexIMG")=request.form("indexIMG")
end if
rs.Update 
rs.Close
if Request("boardmaster")<>"" then call addmaster(Request("boardmaster"),"none",0)
if request("class")<>"0" then
if depth>0 then
	'���ϼ�������ȴ���0��ʱ��Ҫ�����丸�ࣨ����ĸ��ࣩ�İ��������������
	for i=1 to depth
		'�����丸�������
		if parentid<>"" then
		conn.execute("update board set child=child+1 where boardid="&parentid)
		end if
		'�õ��丸��ĸ���İ���ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
		end if
		'��ѭ����������1�������е����һ��ѭ����ʱ��ֱ�ӽ��и���
		if i=depth and parentid<>"" then
		conn.execute("update board set child=child+1 where boardid="&parentid)
		end if
	next
	'���¸ð��������Լ����ڱ���Ҫ��ͬ�ڱ������µİ����������
	conn.execute("update board set orders=orders+1 where rootid="&rootid&" and orders>"&orders)
	conn.execute("update board set orders="&orders&"+1 where boardid="&Request.form("newboardid"))
else
	'���ϼ��������Ϊ0��ʱ��ֻҪ�����ϼ�����������͸ð���������ż���
	conn.execute("update board set child=child+1 where boardid="&request("class"))
	set rs=conn.execute("select max(orders) from board where boardid="&Request.form("newboardid"))
	conn.execute("update board set orders="&rs(0)&"+1 where boardid="&Request.form("newboardid"))
end if
end if
response.write "<p>��̳��ӳɹ���<br><B>����̳Ŀǰ�߼�����ΪĬ��ѡ�������������̳���������������ø���̳�ĸ߼�ѡ��</B><BR><br>"&str
set rs=nothing
call cache_board()
end sub

'����༭��̳��Ϣ
sub savedit()
if clng(request("editid"))=clng(request("class")) then
	response.write "������̳����ָ���Լ�"
	exit sub
end if
dim newboardid,maxrootid
dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr
dim trs,brs,mrs
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="&request("editid")
rs.Open sql,conn,1,3
newboardid=rs("boardid")
parentid=rs("parentid")
iparentid=rs("parentid")
boardmaster=rs("boardmaster")
ParentStr=rs("ParentStr")
depth=rs("depth")
child=rs("child")
rootid=rs("rootid")
'�ж���ָ������̳�Ƿ���������̳
if ParentID=0 then
	if clng(request("class"))<>0 then
	set trs=conn.execute("select rootid from board where boardid="&request("class"))
	if rootid=trs(0) then
		response.write "������ָ���ð����������̳��Ϊ������̳"
		exit sub
	end if
	end if
else
	set trs=conn.execute("select boardid from board where (ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%') and boardid="&request("class"))
	if not (trs.eof and trs.bof) then
		response.write "������ָ���ð����������̳��Ϊ������̳"
		response.end
	end if
end if
if parentid=0 then
	parentid=rs("boardid")
	iparentid=0
end if
rs("boardtype") = Request.Form("boardtype")
'rs("parentid") = Request.Form("class")
rs("boardmaster") = Request("boardmaster")
rs("readme") = Request("readme")
rs("indexIMG")=request.form("indexIMG")
rs("sid")=request.form("sid")
rs.Update 
rs.Close
set rs=nothing
if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)

set mrs=conn.execute("select max(rootid) from board")
Maxrootid=mrs(0)+1
'���������������̳
'��Ҫ������ԭ������������Ϣ��������ȡ�����ID�������������򡢼̳а���������
'��Ҫ���µ�ǰ����������Ϣ
'�̳а���������Ҫ��д�������и���--ȡ������ǰ̨����boardid in parentstr�����
dim k,nParentStr,mParentStr
dim ParentSql,boardcount
if clng(parentid)<>clng(request("class")) and not (iparentid=0 and cint(request("class"))=0) then
	'���ԭ������һ������ĳ�һ������
	if iparentid>0 and cint(request("class"))=0 then
		'���µ�ǰ��������
		conn.execute("update board set depth=0,orders=0,rootid="&maxrootid&",parentid=0,parentstr='0' where boardid="&newboardid)
		ParentStr=ParentStr & ","
		set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
		boardcount=rs(0)
		if isnull(boardcount) then
		boardcount=1
		else
		boardcount=boardcount+1
		end if
		'������ԭ��������̳������
		conn.execute("update board set child=child-"&boardcount&" where boardid="&iparentid)
		'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
		for i=1 to depth
			'�õ��丸��ĸ���İ���ID
			set rs=conn.execute("select parentid from board where boardid="&iparentid)
			if not (rs.eof and rs.bof) then
				iparentid=rs(0)
				conn.execute("update board set child=child-"&boardcount&" where boardid="&iparentid)
			end if
		next
		if child>0 then
		'������������̳����
		'��������̳�������迼�ǣ�����������̳��Ⱥ�һ������ID(rootid)����
		'���µ�ǰ��������
		'ParentStr=ParentStr & ","		
		i=0
		set rs=conn.execute("select * from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
		do while not rs.eof
		i=i+1
		mParentStr=replace(rs("ParentStr"),ParentStr,"")
		conn.execute("update board set depth=depth-"&depth&",rootid="&maxrootid&",ParentStr='"&mParentStr&"' where boardid="&rs("boardid"))
		rs.movenext
		loop
		end if
	elseif iparentid>0 and cint(request("class"))>0 then
	'��һ������̳�ƶ�����������̳��
	'�����ָ������̳�������Ϣ
	set trs=conn.execute("select * from board where boardid="&request("class"))
	'�õ�������������
	ParentStr=ParentStr & ","
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	boardcount=rs(0)
	if isnull(boardcount) then boardcount=1
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	conn.execute("update board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	'���µ�ǰ��������
	conn.execute("update board set depth="&trs("depth")&"+1,orders="&trs("orders")&"+1,rootid="&trs("rootid")&",ParentID="&request("class")&",ParentStr='" & trs("parentstr") & "," & trs("boardid") & "' where boardid="&newboardid)
	i=1
	'����������������������
	'���Ϊԭ����ȼ��ϵ�ǰ������̳�����
	set rs=conn.execute("select * from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	iParentStr=trs("parentstr") & "," & trs("boardid") & "," & replace(rs("parentstr"),ParentStr,"")
	conn.execute("update board set depth=depth+"&trs("depth")&"-"&depth&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&iParentStr&"' where boardid="&rs("boardid"))
	rs.movenext
	loop
	ParentID=request("class")
	if rootid=trs("rootid") then
	'��ͬһ�������ƶ�
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
	conn.execute("update board set child=child+"&i&" where (not ParentID=0) and boardid="&parentid)
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=conn.execute("select parentid from board where (not ParentID=0) and boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			conn.execute("update board set child=child+"&i&" where (not ParentID=0) and  boardid="&parentid)
		end if
	next
	'������ԭ���������
	conn.execute("update board set child=child-"&i&" where (not ParentID=0) and boardid="&iparentid)
	'������ԭ��������̳����
	'response.write iparentid & "<br>"
	for k=1 to depth
		'�õ���ԭ����ĸ���İ���ID
		set rs=conn.execute("select parentid from board where (not ParentID=0) and boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'response.write iparentid & "<br>"
			'������ԭ����ĸ��������
			conn.execute("update board set child=child-"&i&" where (not ParentID=0) and  boardid="&iparentid)
		end if
	next
	else
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
	conn.execute("update board set child=child+"&i&" where boardid="&parentid)
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			conn.execute("update board set child=child+"&i&" where boardid="&parentid)
		end if
	next
	'������ԭ���������
	conn.execute("update board set child=child-"&i&" where boardid="&iparentid)
	'������ԭ��������̳����
	for k=1 to depth
		'�õ���ԭ����ĸ���İ���ID
		set rs=conn.execute("select parentid from board where boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'������ԭ����ĸ��������
			conn.execute("update board set child=child-"&i&" where boardid="&iparentid)
		end if
	next
	end if 'end if rootid=trs("rootid") then
	else
	'���ԭ����һ����̳�ĳ�������̳��������̳
	'�õ���ָ������̳�������Ϣ
	set trs=conn.execute("select * from board where boardid="&request("class"))
	set rs=conn.execute("select count(*) from board where rootid="&rootid)
	boardcount=rs(0)
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	ParentID=request("class")
	'�����丸�������
	conn.execute("update board set child=child+"&boardcount&" where boardid="&parentid)
	'response.write parentid & "-"&boardcount&"<br>"
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=conn.execute("select parentid from board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			conn.execute("update board set child=child+"&boardcount&" where boardid="&parentid)
		end if
	'response.write parentid & "-"&boardcount&"<br>"
	next
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	conn.execute("update board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	i=0
	set rs=conn.execute("select * from board where rootid="&rootid&" order by orders")
	do while not rs.eof
	i=i+1
	if rs("parentid")=0 then
		if trs("ParentStr")="0" then
		parentstr=trs("boardid")
		else
		parentstr=trs("parentstr") & "," & trs("boardid")
		end if
	conn.execute("update board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"',parentid="&request("class")&" where boardid="&rs("boardid"))
	else
		if trs("ParentStr")="0" then
		parentstr=trs("boardid") & "," & rs("parentstr")
		else
		parentstr=trs("parentstr") & "," & trs("boardid") & "," & rs("parentstr")
		end if
	conn.execute("update board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"' where boardid="&rs("boardid"))
	end if
	rs.movenext
	loop
	end if
end if
response.write "<p>��̳�޸ĳɹ���<br><br>"&str
set rs=nothing
set mrs=nothing
set trs=nothing
'cache��������
call cache_board()
'end cache
end sub

'ɾ�����棬ɾ���������ӣ���ڣ�����ID
sub del()
'�������ϼ�������̳�����������̳�����¼���̳������ɾ��
set rs=conn.execute("select ParentStr,child,depth from board where boardid="&Request("editid"))
if not (rs.eof and rs.bof) then
if rs(1)>0 then
	response.write "����̳����������̳����ɾ����������̳���ٽ���ɾ������̳�Ĳ���"
	exit sub
end if
'������ϼ����棬���������
if rs(2)>0 then
	conn.execute("update board set child=child-1 where boardid in ("&rs(0)&")")
end if
sql = "delete from board where boardid="&Request("editid")
conn.execute(sql)
for i=0 to ubound(AllPostTable)
sql = "delete from "&AllPostTable(i)&" where boardid="&Request("editid")
conn.execute(sql)
next
end if
set rs=nothing
call cache_board()
response.write "<p>��̳ɾ���ɹ���"
end sub

sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳һ���������������޸�(������Ӧ��̳������������������Ӧ���������)
	</th>
	</tr>

	<tr>
	<td class="Forumrow"><table width="50%">
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from Board where ParentID=0 order by RootID"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "��û����Ӧ����̳���ࡣ"
	else
		do while not rs.eof
		response.write "<form action=admin_board.asp?action=updatorders method=post><tr><td width=""50%"">"&rs("boardtype")&"</td>"
		response.write "<td width=""50%""><input type=text name=""OrderID"" size=4 value="""&rs("rootid")&"""><input type=hidden name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=submit name=Submit value=�޸�></td></tr></form>"
		rs.movenext
		loop
%>
</table>
<BR>&nbsp;<font color=red>��ע�⣬����һ��<B>������д��ͬ�����</B>������ǳ����޸���</font>
<%
	end if
	rs.close
	set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateorders()
	dim cID,OrderID,ClassName
	'response.write request.form("cID")(1)
	'response.end
	cID=replace(request.form("cID"),"'","")
	OrderID=replace(request.form("OrderID"),"'","")
	set rs=conn.execute("select boardid from board where rootid="&orderid)
	if rs.eof and rs.bof then
	response.write "���óɹ����뷵�ء�"
	conn.execute("update board set rootid="&OrderID&" where rootid="&cID)
	else
	response.write "�벻Ҫ��������̳������ͬ�����"
	end if
	call cache_board()
end sub


sub boardorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳N���������������޸�(������Ӧ��̳������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
dim trs,uporders,doorders
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from Board order by RootID,orders"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "��û����Ӧ����̳���ࡣ"
else
	do while not rs.eof
	response.write "<form action=admin_board.asp?action=updatboardorders method=post><tr><td width=""50%"">"
	if rs("depth")>0 then
	for i=1 to rs("depth")
		response.write "&nbsp;"
	next
	end if
	if rs("child")>0 then
		response.write "<img src=pic/plus.gif>"
	else
		response.write "<img src=pic/nofollow.gif>"
	end if
	if rs("parentid")=0 then
		response.write "<b>"
	end if
	response.write rs("boardtype")
	if rs("child")>0 then
		response.write "("&rs("child")&")"
	end if
	response.write "</td><td width=""50%"">"
	if rs("ParentID")>0 then
	'�����ͬ��ȵİ�����Ŀ���õ��ð�������ͬ��ȵİ���������λ�ã�֮�ϻ���֮�µİ�������
	'��������������ӦΪFor i=1 to �ð�֮�ϵİ�����
	set trs=conn.execute("select count(*) from board where ParentID="&rs("ParentID")&" and orders<"&rs("orders")&"")
	uporders=trs(0)
	if isnull(uporders) then uporders=0
	'���ܽ���������ӦΪFor i=1 to �ð�֮�µİ�����
	set trs=conn.execute("select count(*) from board where ParentID="&rs("ParentID")&" and orders>"&rs("orders")&"")
	doorders=trs(0)
	if isnull(doorders) then doorders=0
	if uporders>0 then
		response.write "<select name=uporders size=1><option value=0>�����ƶ�</option>"
		for i=1 to uporders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 then
		if uporders>0 then response.write "&nbsp;"
		response.write "<select name=doorders size=1><option value=0>�����ƶ�</option>"
		for i=1 to doorders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 or uporders>0 then
	response.write "<input type=hidden name=""editID"" value="""&rs("boardid")&""">&nbsp;<input type=submit name=Submit value=�޸�>"
	end if
	end if
	response.write "</td></tr></form>"
	uporders=0
	doorders=0
	rs.movenext
	loop
	response.write "</table>"
end if
rs.close
set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateboardorders()
dim ParentID,orders,ParentStr,child
dim uporders,doorders,oldorders,trs,ii
if not isnumeric(request("editID")) then
	response.write "�Ƿ��Ĳ�����"
	exit sub
end if
if request("uporders")<>"" and not Cint(request("uporders"))=0 then
	if not isnumeric(request("uporders")) then
	response.write "�Ƿ��Ĳ�����"
	exit sub
	elseif Cint(request("uporders"))=0 then
	response.write "��ѡ��Ҫ���������֣�"
	exit sub
	end if
	'�����ƶ�
	'Ҫ�ƶ�����̳��Ϣ
	set rs=conn.execute("select ParentID,orders,ParentStr,child from board where boardid="&request("editID"))
	ParentID=rs(0)
	orders=rs(1)
	ParentStr=rs(2) & "," & request("editID")
	child=rs(3)
	i=0
	'response.write "select boardid,orders from board where ParentID="&ParentID&" and orders<"&orders&" order by orders desc<br>"
	if child>0 then
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	'�͸���̳ͬ������������֮�ϵ���̳��������������ĩ��Ϊ��ǰ��̳�����
	set rs=conn.execute("select boardid,orders,child,ParentStr from board where ParentID="&ParentID&" and orders<"&orders&" order by orders desc")
	do while not rs.eof
	i=i+1
	if Cint(request("uporders"))>=i then
		'response.write "update board set orders="&orders&" where boardid="&rs(0)&"<br>"
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select boardid,orders from board where ParentStr='"&rs(3)&","&rs(0)&"' or ParentStr like '"&rs(3)&","&rs(0)&",%' order by orders")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		conn.execute("update board set orders="&orders&"+"&oldorders&"+"&ii&" where boardid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		conn.execute("update board set orders="&orders&"+"&oldorders&" where boardid="&rs(0))
		if Cint(request("uporders"))=i then uporders=rs(1)
	end if
	orders=rs(1)
	rs.movenext
	loop
	'response.write "update board set orders="&uporders&" where boardid="&request("editID")
	'������Ҫ�������̳�����
	conn.execute("update board set orders="&uporders&" where boardid="&request("editID"))
	'�����������̳���������������̳����
	if child>0 then
	i=uporders
	set rs=conn.execute("select boardid from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	conn.execute("update board set orders="&i&" where boardid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
elseif request("doorders")<>"" then
	if not isnumeric(request("doorders")) then
	response.write "�Ƿ��Ĳ�����"
	exit sub
	elseif Cint(request("doorders"))=0 then
	response.write "��ѡ��Ҫ�½������֣�"
	exit sub
	end if
	set rs=conn.execute("select ParentID,orders,ParentStr,child from board where boardid="&request("editID"))
	ParentID=rs(0)
	orders=rs(1)
	ParentStr=rs(2) & "," & request("editID")
	child=rs(3)
	i=0
	if child>0 then
	set rs=conn.execute("select count(*) from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	set rs=conn.execute("select boardid,orders,child,ParentStr from board where ParentID="&ParentID&" and orders>"&orders&" order by orders")
	do while not rs.eof
	i=i+1
	if Cint(request("doorders"))>=i then
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select boardid,orders from board where ParentStr='%"&rs(3)&","&rs(0)&"%' or ParentStr like '"&rs(3)&","&rs(0)&",%' order by orders")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		'response.write "update board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
		conn.execute("update board set orders="&orders&"+"&ii&" where boardid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		'response.write "update board set orders="&orders&" where boardid="&rs(0)&"<br>"
		conn.execute("update board set orders="&orders&" where boardid="&rs(0))
		if Cint(request("doorders"))=i then doorders=rs(1)
	end if
	orders=rs(1)
	rs.movenext
	loop
	'response.write "update board set orders="&doorders&" where boardid="&request("editID")&"<br>"
	conn.execute("update board set orders="&doorders&" where boardid="&request("editID"))
	'�����������̳���������������̳����
	if child>0 then
	i=doorders
	set rs=conn.execute("select boardid from board where ParentStr='"&ParentStr&"' or ParentStr like '"&ParentStr&",%' order by orders")
	do while not rs.eof
	i=i+1
	'response.write "update board set orders="&i&" where boardid="&rs(0)&"��b<br>"
	conn.execute("update board set orders="&i&" where boardid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
end if
call cache_board()
response.redirect "admin_board.asp?action=boardorders"
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

sub boardpermission()
dim iUserGroupID(20),UserTitle(20)
dim trs,ars,k,ii
ii=0
set trs=conn.execute("select title,usergroupid from usergroups order by usergroupid")
do while not trs.eof
UserTitle(ii)=trs(0)
iUserGroupID(ii)=trs(1)
ii=ii+1
trs.movenext
loop
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height="25">�༭��̳Ȩ��</th>
	</tr>
	<tr>
	<td class=forumrow>�����������ò�ͬ�û����ڲ�ͬ��̳�ڵ�Ȩ�ޣ���ɫ��ʾΪ����̳���û���ʹ�õ����û���������<BR>�ڸ�Ȩ�޲��ܼ̳У�������������һ�������¼���̳�İ��棬��ôֻ�������õİ�����Ч��������������̳��Ч<BR>���������������Ч������������ҳ��<B>ѡ���Զ�������</B>��ѡ�����Զ������ú��������õ�Ȩ�޽�<B>����</B>���û������ã������û���Ĭ�ϲ��ܹ������ӣ������������˸��û���ɹ������ӣ���ô���û������������Ϳ��Թ�������
	</td>
	</tr>
</table><BR>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>�����û���Ȩ��
</th>
</tr>
<%
sql="select * from board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
%>
<tr> 
<td height="25" width=40%  class="forumrow">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<FORM METHOD=POST ACTION="?action=editpermission">
<td width=60% class="forumrow">&nbsp;
<select name="groupid" size=1>
<%
	for k=0 to ii-1
		set ars=conn.execute("select pid from BoardPermission where BoardID="&rs("boardid")&" and GroupID="&iUserGroupID(k))
		if ars.eof and ars.bof then
		response.write "<option value="""&iUserGroupID(k)&""">" & UserTitle(k) & "</option>"
		else
		response.write "<option value="""&iUserGroupID(k)&""">" & UserTitle(k) & "(�Զ���)</option>"
		end if
	next
%>
</select>
<input type=hidden value="<%=rs("boardid")%>" name=reboardid>
<input type=submit name=submit value="����">
<%
	dim percount
	set trs=conn.execute("select count(*) from BoardPermission where boardid="&rs("boardid"))
	percount=trs(0)
	if not isnull(percount) and percount>0 then response.write "(���Զ������)"

%>
</td>
</FORM>
</tr>
<%
rs.movenext
loop
set rs=nothing
%>
</table><BR><BR>
<%
set trs=nothing
set ars=nothing
end sub

sub editpermission()
if not isnumeric(request("groupid")) then
response.write "����Ĳ�����"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	GroupSetting=Request.Form("canview") & "," & Request.Form("canviewuserinfo") & "," & Request.Form("canviewpost") & "," & Request.Form("cannewpost") & "," & Request.Form("canreplymytopic") & "," & Request.Form("canreplytopic") & "," & Request.Form("canpostagree") & "," & Request.Form("canupload") & "," & Request.Form("canpostvote") & "," & Request.Form("canvote") & "," & Request.Form("caneditmytopic") & "," & Request.Form("candelmytopic") & "," & Request.Form("canmovemytopic") & "," & Request.Form("canclosemytopic") & "," & Request.Form("cansearch") & "," & Request.Form("canmailtopic") & "," & Request.Form("canmodify") & "," & Request.Form("cansmallpaper") & "," & Request.Form("candeltopic") & "," & Request.Form("canmovetopic") & "," & Request.Form("canclosetopic") & "," & Request.Form("cantoptopic") & "," & Request.Form("canawardtopic") & "," & Request.Form("canmodifytopic") & "," & Request.Form("canbesttopic") & "," & Request.Form("canAnnounce") & "," & Request.Form("canAdminAnnounce") & "," & Request.Form("canAdminPaper") & "," & Request.Form("canAdminUser") & "," & Request.Form("canDelUserTopic") & "," & Request.Form("canviewip") & "," & Request.Form("canadminip") & "," & Request.Form("cansendsms") & "," & Request.Form("Maxsendsms") & "," & Request.Form("Maxsmsbody") & "," & Request.Form("Maxsmsbox") & "," & Request.Form("canusetitle") & "," & Request.Form("canuseface") & "," & Request.Form("canusesign") & "," & Request.Form("canvieweven") & "," & Request.Form("canuploadnum") & "," & Request.Form("canviewbest") & "," & Request.Form("adminpermission") & "," & request.form("canaward") & "," & request.form("MaxUploadSize") & "," & request.form("canbatchtopic") & "," & request.form("smallpapermoney") & "," & request.form("postagreemoney") & "," & request.form("canadminfile") & "," & request.form("ba1") & "," & Request.Form("ba2") & "," & request.form("ba3") & "," & request.form("ba4") & "," & request.form("ba5") & "," & request.form("ba6") & "," & request.form("ba7")
	Set rs= Server.CreateObject("ADODB.Recordset")
	if request("isdefault")=1 then
	conn.execute("delete from BoardPermission where BoardID="&request("reBoardID")&" and GroupID="&request("GroupID"))
	else
	if request("pid")<>"" then
	sql="update BoardPermission set PSetting='"&GroupSetting&"' where pid="&request("pid")
	else
	sql="insert into BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&request("GroupID")&",'"&GroupSetting&"')"
	end if
	conn.execute(sql)
	end if
	set rs=nothing
	response.write "�޸ĳɹ�������<a href=?action=permission>��̳Ȩ�޹���</a>"
else
Dim reGroupSetting,reBoardID,groupid
Dim Groupname,Boardname,founduserper
founduserper=false
if request("GroupID")<>"" then
set rs=conn.execute("select * from BoardPermission where boardid="&request("reBoardID")&" and GroupID="&request("GroupID"))
if rs.eof and rs.bof then
	founduserper=false
else
groupid=rs("groupid")
reGroupSetting=split(rs("PSetting"),",")
reBoardID=rs("boardid")
set rs=conn.execute("select title from UserGroups where usergroupid="&groupid)
groupname=rs("title")
founduserper=true
end if
if not founduserper then
set rs=conn.execute("select * from usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "δ�ҵ����û��飡"
exit sub
end if
groupid=request("groupid")
reGroupSetting=split(rs("GroupSetting"),",")
reBoardID=request("reBoardID")
Groupname=rs("title")
end if
end if
set rs=conn.execute("select boardtype from board where boardid="&reBoardID)
Boardname=rs("boardtype")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=editpermission">
<input type=hidden name="groupid" value="<%=groupid%>">
<input type=hidden name="reBoardID" value="<%=reBoardID%>">
<input type=hidden name="pID" value="<%
	dim ars
	set ars=conn.execute("select pid from BoardPermission where BoardID="&reBoardID&" and GroupID="&groupid)
	if not (ars.eof or ars.bof) then
		response.write ars("pid")
	end if
%>">
<tr> 
<th height="23" colspan="2" >�༭��̳�û���Ȩ��&nbsp;>> <%=boardname%>&nbsp;>> <%=groupname%></th>
</tr>
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>ʹ���Զ�������</B>&nbsp;(ѡ���Զ������ʹ����������Ч) </td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>�����鿴Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���������̳</td>
<td height="23" width="40%" class=forumrow>��<input name="canview" type=radio value="1" <%if reGroupSetting(0)=1 then%>checked<%end if%>>&nbsp;��<input name="canview" type=radio value="0" <%if reGroupSetting(0)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=forumrow>��<input name="canviewuserinfo" type=radio value="1" <%if reGroupSetting(1)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewuserinfo" type=radio value="0" <%if reGroupSetting(1)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canviewpost" type=radio value="1" <%if reGroupSetting(2)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewpost" type=radio value="0" <%if reGroupSetting(2)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="canviewbest" type=radio value="1" <%if reGroupSetting(41)=1 then%>checked<%end if%>>&nbsp;��<input name="canviewbest" type=radio value="0" <%if reGroupSetting(41)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Է���������</td>
<td height="23" width="40%" class=forumrow>��<input name="cannewpost" type=radio value="1" <%if reGroupSetting(3)=1 then%>checked<%end if%>>&nbsp;��<input name="cannewpost" type=radio value="0" <%if reGroupSetting(3)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canreplymytopic" type=radio value="1" <%if reGroupSetting(4)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplymytopic" type=radio value="0" <%if reGroupSetting(4)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Իظ������˵�����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canreplytopic" type=radio value="1" <%if reGroupSetting(5)=1 then%>checked<%end if%>>&nbsp;��<input name="canreplytopic" type=radio value="0" <%if reGroupSetting(5)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=forumrow>��<input name="canpostagree" type=radio value="1" <%if reGroupSetting(6)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostagree" type=radio value="0" <%if reGroupSetting(6)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������������Ǯ
</td>
<td height="23" width="40%" class=Forumrow><input name="postagreemoney" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����ϴ�����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canupload" type=radio value="1" <%if reGroupSetting(7)=1 then%>checked<%end if%>>&nbsp;��<input name="canupload" type=radio value="0" <%if reGroupSetting(7)=0 then%>checked<%end if%>>
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
<td height="23" width="60%" class=forumrow>���Է�����ͶƱ</td>
<td height="23" width="40%" class=forumrow>��<input name="canpostvote" type=radio value="1" <%if reGroupSetting(8)=1 then%>checked<%end if%>>&nbsp;��<input name="canpostvote" type=radio value="0" <%if reGroupSetting(8)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Բ���ͶƱ</td>
<td height="23" width="40%" class=forumrow>��<input name="canvote" type=radio value="1" <%if reGroupSetting(9)=1 then%>checked<%end if%>>&nbsp;��<input name="canvote" type=radio value="0" <%if reGroupSetting(9)=0 then%>checked<%end if%>></td>
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
<th height="23" colspan="2" align=left>����<b>����/����༭Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="caneditmytopic" type=radio value="1" <%if reGroupSetting(10)=1 then%>checked<%end if%>>&nbsp;��<input name="caneditmytopic" type=radio value="0" <%if reGroupSetting(10)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="candelmytopic" type=radio value="1" <%if reGroupSetting(11)=1 then%>checked<%end if%>>&nbsp;��<input name="candelmytopic" type=radio value="0" <%if reGroupSetting(11)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmovemytopic" type=radio value="1" <%if reGroupSetting(12)=1 then%>checked<%end if%>>&nbsp;��<input name="canmovemytopic" type=radio value="0" <%if reGroupSetting(12)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canclosemytopic" type=radio value="1" <%if reGroupSetting(13)=1 then%>checked<%end if%>>&nbsp;��<input name="canclosemytopic" type=radio value="0" <%if reGroupSetting(13)=0 then%>checked<%end if%>></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>����<b>����Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����������̳
</td>
<td height="23" width="40%" class=forumrow>��<input name="cansearch" type=radio value="1" <%if reGroupSetting(14)=1 then%>checked<%end if%>>&nbsp;��<input name="cansearch" type=radio value="0" <%if reGroupSetting(14)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmailtopic" type=radio value="1" <%if reGroupSetting(15)=1 then%>checked<%end if%>>&nbsp;��<input name="canmailtopic" type=radio value="0" <%if reGroupSetting(15)=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=forumrow>�����޸ĸ�������
</td>
<td height="23" width="40%" class=forumrow>��<input name="canmodify" type=radio value="1" <%if reGroupSetting(16)=1 then%>checked<%end if%>>&nbsp;��<input name="canmodify" type=radio value="0" <%if reGroupSetting(16)=0 then%>checked<%end if%>></td>
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
<tr> 
<td height="23" width="60%" class=forumrow>
</td>
<td height="23" width="40%" class=forumrow><input type="submit" name="submit" value="�� ��"></td>
</tr>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub cache_board()
'cache��������
myCache.name="BoardJumpList"
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
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
myCache.add BoardJumpList,dateadd("n",9999,now)
set rs=nothing
'end cache
end sub

sub RestoreBoard()
'����Ŀǰ������ѭ��i��ֵ����rootid
'��ԭ���а����depth,orders,parentid,parentstr,childΪ0
i=0
set rs=conn.execute("select boardid from board order by rootid,orders")
do while not rs.eof
i=i+1
conn.execute("update board set rootid="&i&",depth=0,orders=0,ParentID=0,ParentStr='0',child=0 where boardid="&rs(0))
rs.movenext
loop
set rs=nothing
response.write "��λ�ɹ����뷵������̳�������á�"
call cache_board()
end sub
%>