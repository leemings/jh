<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	dim admin_flag
	admin_flag="23"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
		call savegrade()
		else
		call grade()
		end if
		if founderr then call dvbbs_error()
		conn.close
		set conn=nothing
	end if

sub grade()
dim sel
%>
<form method="POST" action=admin_wealth.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="2" ><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳���ģ����<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳�����У��ɶ�ѡ<BR>3�����������һ���������ñ�İ����ģ������ã�ֻҪ����ð����ģ�����ƣ������ʱ��ѡ��Ҫ���浽�İ������ƻ�ģ�����Ƽ��ɡ�</font></td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2><B>��ǰʹ����ģ��</B>���ɽ����ñ��浽����ģ���У�</th>
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
	skinid=rs("id")
	Forum_user=split(rs("Forum_user"),",")
	uForum_user=ubound(Forum_user)
	if uForum_user<39 then
		redim preserve Forum_user(39)
		Forum_user(18)=100
		Forum_user(19)=50
		Forum_user(20)=50
		Forum_user(21)=25
		Forum_user(22)=20
		Forum_user(23)=10
		Forum_user(24)=40
		Forum_user(25)=20
		Forum_user(26)=4
		Forum_user(27)=2
		Forum_user(28)=40
		Forum_user(29)=20
		Forum_user(30)=0
		Forum_user(31)=0
		Forum_user(32)=10
		Forum_user(33)=5
		Forum_user(34)=2
		Forum_user(35)=1
		Forum_user(36)=2
		Forum_user(37)=1
		Forum_user(38)=2
		Forum_user(39)=1
	end if
	if uForum_user<42 then
		redim preserve Forum_user(42)
		Forum_user(40)=5
		Forum_user(41)=1
		Forum_user(42)=1
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
	Forum_user=split(rs("Forum_user"),",")
	uForum_user=ubound(Forum_user)
	if uForum_user<39 then
		redim preserve Forum_user(39)
		Forum_user(18)=100
		Forum_user(19)=50
		Forum_user(20)=50
		Forum_user(21)=25
		Forum_user(22)=20
		Forum_user(23)=10
		Forum_user(24)=40
		Forum_user(25)=20
		Forum_user(26)=4
		Forum_user(27)=2
		Forum_user(28)=40
		Forum_user(29)=20
		Forum_user(30)=0
		Forum_user(31)=0
		Forum_user(32)=10
		Forum_user(33)=5
		Forum_user(34)=2
		Forum_user(35)=1
		Forum_user(36)=2
		Forum_user(37)=1
		Forum_user(38)=2
		Forum_user(39)=1
	end if
	if uForum_user<42 then
		redim preserve Forum_user(42)
		Forum_user(40)=5
		Forum_user(41)=1
		Forum_user(42)=1
	end if
	else
	sel=""
	end if
end if
response.write "<input type=checkbox name=skinid value="&rs("id")&" "&sel&"><a href=admin_wealth.asp?skinid="&rs("id")&">"&rs("skinname")&"</a>&nbsp;"
rs.movenext
loop
rs.close
set rs=nothing
%>
</td></tr>
<tr> 
<td width="100%" colspan=2 class=Forumrow><B>��ǰʹ�ø�ģ��ķ���̳</B><BR>���·���̳��ʹ�ñ�ģ�����ã������޸���̳ʹ��ģ�壬�ɵ���̳��������������<BR>
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
<th height="23" colspan="2">�û���Ǯ�趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע���Ǯ��</td>
<td width="60%" class=Forumrow><input type="text" name="wealthReg" size="35" value="<%=Forum_user(0)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼���ӽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="wealthLogin" size="35" value="<%=Forum_user(4)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="wealthAnnounce" size="35" value="<%=Forum_user(1)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="wealthReannounce" size="35" value="<%=Forum_user(2)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="BestWealth" size="35" value="<%=Forum_user(15)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����ٽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="wealthDel" size="35" value="<%=Forum_user(3)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ͶƱ���ӽ�Ǯ</td>
<td width="60%" class=Forumrow><input type="text" name="wealthVote" size="35" value="<%=Forum_user(40)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע�ᾭ��ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epReg" size="35" value="<%=Forum_user(5)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼���Ӿ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epLogin" size="35" value="<%=Forum_user(9)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epAnnounce" size="35" value="<%=Forum_user(6)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epReannounce" size="35" value="<%=Forum_user(7)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="bestuserep" size="35" value="<%=Forum_user(17)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����پ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epDel" size="35" value="<%=Forum_user(8)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ͶƱ���Ӿ���ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="epVote" size="35" value="<%=Forum_user(41)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpReg" size="35" value="<%=Forum_user(10)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼��������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpLogin" size="35" value="<%=Forum_user(14)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpAnnounce" size="35" value="<%=Forum_user(11)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpReannounce" size="35" value="<%=Forum_user(12)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="bestusercp" size="35" value="<%=Forum_user(16)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ����������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpDel" size="35" value="<%=Forum_user(13)%>"></td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ͶƱ��������ֵ</td>
<td width="60%" class=Forumrow><input type="text" name="cpVote" size="35" value="<%=Forum_user(42)%>"></td>
</tr>
<tr> 
<th height="23" colspan="2" >�����շѼ۸��趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա��ȫ���Ա���</td>
<td width="60%" class=Forumrow><input type="text" name="dgallDel" size="35" value="<%=Forum_user(18)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա����һ��Ա���</td>
<td width="60%" class=Forumrow><input type="text" name="dgDel" size="35" value="<%=Forum_user(19)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա��ȫ���Ա���</td>
<td width="60%" class=Forumrow><input type="text" name="dgvipallDel" size="35" value="<%=Forum_user(20)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա����һ��Ա���</td>
<td width="60%" class=Forumrow><input type="text" name="dgvipDel" size="35" value="<%=Forum_user(21)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�޸�һ��ͷ��</td>
<td width="60%" class=Forumrow><input type="text" name="faceDel" size="35" value="<%=Forum_user(22)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�޸�һ��ͷ��</td>
<td width="60%" class=Forumrow><input type="text" name="facevipDel" size="35" value="<%=Forum_user(23)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�޸�һ��ǩ��</td>
<td width="60%" class=Forumrow><input type="text" name="signDel" size="35" value="<%=Forum_user(24)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�޸�һ��ǩ��</td>
<td width="60%" class=Forumrow><input type="text" name="signvipDel" size="35" value="<%=Forum_user(25)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա���Ͷ���</td>
<td width="60%" class=Forumrow><input type="text" name="mesDel" size="35" value="<%=Forum_user(26)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա���Ͷ���</td>
<td width="60%" class=Forumrow><input type="text" name="mesvipDel" size="35" value="<%=Forum_user(27)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�ظ������֪ͨ</td>
<td width="60%" class=Forumrow><input type="text" name="msgDel" size="35" value="<%=Forum_user(28)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�ظ������֪ͨ</td>
<td width="60%" class=Forumrow><input type="text" name="msgvipDel" size="35" value="<%=Forum_user(29)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�ظ����ʼ�֪ͨ</td>
<td width="60%" class=Forumrow><input type="text" name="emailDel" size="35" value="<%=Forum_user(30)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�ظ����ʼ�֪ͨ</td>
<td width="60%" class=Forumrow><input type="text" name="emailvipDel" size="35" value="<%=Forum_user(31)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա��������ע��</td>
<td width="60%" class=Forumrow><input type="text" name="notifyDel" size="35" value="<%=Forum_user(32)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա��������ע��</td>
<td width="60%" class=Forumrow><input type="text" name="notifyvipDel" size="35" value="<%=Forum_user(33)%>">&nbsp;Ԫ/��</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�ϴ�ͷ��</td>
<td width="60%" class=Forumrow><input type="text" name="upfaceDel" size="35" value="<%=Forum_user(34)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�ϴ�ͷ��</td>
<td width="60%" class=Forumrow><input type="text" name="upfacevipDel" size="35" value="<%=Forum_user(35)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�ϴ��ļ�</td>
<td width="60%" class=Forumrow><input type="text" name="upfileDel" size="35" value="<%=Forum_user(36)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�ϴ��ļ�</td>
<td width="60%" class=Forumrow><input type="text" name="upfilevipDel" size="35" value="<%=Forum_user(37)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��ͨ��Ա�ϴ���Ƭ</td>
<td width="60%" class=Forumrow><input type="text" name="upphotoDel" size="35" value="<%=Forum_user(38)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>VIP��Ա�ϴ���Ƭ</td>
<td width="60%" class=Forumrow><input type="text" name="upphotovipDel" size="35" value="<%=Forum_user(39)%>">&nbsp;Ԫ/KB</td>
</tr>
<tr height=23> 
<td colspan=2 class=Forumrow align=center><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
if trim(request("skinid"))="" and trim(request("boardid"))="" then
Founderr=true
Errmsg=Errmsg+"<br>"+"<li>��ѡ�񱣴��ģ������"
else
Forum_user=request.form("wealthReg") & "," & request.form("wealthAnnounce") & "," & request.form("wealthReannounce") & "," & request.form("wealthDel") & "," & request.form("wealthLogin") & "," & request.form("epReg") & "," & request.form("epAnnounce") & "," & request.form("epReannounce") & "," & request.form("epDel") & "," & request.form("epLogin") & "," & request.form("cpReg") & "," & request.form("cpAnnounce") & "," & request.form("cpReannounce") & "," & request.form("cpDel") & "," & request.form("cpLogin") & "," & request.form("BestWealth") & "," & request.form("BestuserCP") & "," & request.form("BestuserEP") & "," & request.form("dgalldel") & "," & request.form("dgdel") & "," & request.form("dgvipalldel") & "," & request.form("dgvipdel") & "," & request.form("facedel") & "," & request.form("facevipdel") & "," & request.form("signdel") & "," & request.form("signvipdel") & "," & request.form("mesdel") & "," & request.form("mesvipdel") & "," & request.form("msgdel") & "," & request.form("msgvipdel") & "," & request.form("emaildel") & "," & request.form("emailvipdel") & "," & request.form("notifydel") & "," & request.form("notifyvipdel") & "," & request.form("upfacedel") & "," & request.form("upfacevipdel") & "," & request.form("upfiledel") & "," & request.form("upfilevipdel") & "," & request.form("upphotodel") & "," & request.form("upphotovipdel") & "," & request.form("wealthVote") & "," & request.form("epVote") & "," & request.form("cpVote")
if request("checkall")="1" then
sql="update config set Forum_user='"&Forum_user&"'"
conn.execute(sql)
else
if request("skinid")<>"" then
sql = "update config set Forum_user='"&Forum_user&"' where id in ( "&request("skinid")&" )"
conn.execute(sql)
end if
end if
response.write "���óɹ���"
end if
end sub
%>