<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<script language="JavaScript">
<!--
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<%
	dim admin_flag
	admin_flag="26"
	if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername,oldadduser,username1
'''''''''''''''
'ȡ���û������Ա������ 2002-12-13
		dim groupsname,titlepic
		set rs=conn.execute("select title from [UserGroups] where UserGroupID=1 ")
		groupsname=rs(0)
		set rs=conn.execute("select titlepic from usertitle where usergroupid=1 order by Minarticle desc")
		titlepic=rs(0)
		set rs=nothing

		if request("action")="updat" then
			call update()
			response.write body
		elseif request("action")="del" then
			call del()
			response.write body
        	elseif request("action")="pasword" then
			call pasword()
        	elseif request("action")="newpass" then
			call newpass()
			response.write body
		elseif request("action")="add" then
			call addadmin()
		elseif request("action")="edit" then
			call userinfo()
		elseif request("action")="savenew" then
			call savenew()
			response.write body
		else
			call userlist()
		end if
		conn.close
		set conn=nothing
	end if

	sub userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th height=22 colspan=5>����Ա����(����û������в���)</th>
                </tr>
                <tr align=center> 
                  <td width="30%" height=22 class="forumHeaderBackgroundAlternate"><B>�û���</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>�ϴε�¼ʱ��</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>�ϴε�¼IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>����</B></td>
                </tr>
<%
	set rs=conn.execute("select * from admin order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
                  <td class=forumrow><a href="admin_admin.asp?id=<%=rs("id")%>&action=pasword"><%=rs("username")%></a></td><td class=forumrow><%=rs("LastLogin")%></td><td class=forumrow><%=rs("LastLoginIP")%></td><td class=forumrow><a href="admin_admin.asp?action=del&id=<%=rs("id")%>&name=<%=rs("username")%>">ɾ��</a>&nbsp;&nbsp;<a href="admin_admin.asp?id=<%=rs("id")%>&action=edit">�༭Ȩ��</a></td>
                </tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
	       </table>
<%
	end sub

	sub del()
	conn.execute("delete from admin where id="&request("id"))
	conn.execute("update [user] set usergroupid=4 where username='"&replace(request("name"),"'","")&"'")
	body="<li>����Աɾ���ɹ���"
	end sub

sub pasword()
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
  %> 
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա���Ϲ����������޸�
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" value="<%=oldpassword%>">  (����ע�����벻ͬ,��Ҫ�޸���ֱ������)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><%=oldadduser%>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 

              <input type=hidden name="oldadduser" value="<%=oldadduser%>">
              <input type=hidden name="adduser" value="<%=oldadduser%>">
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

sub newpass()
	dim passnw,usernw,aduser
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		body="<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		usernw=trim(request("username2"))
	end if
	if request("password2")="" then
		body="<li>�������������롣<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"))
	end if
	if request("adduser")="" then
		body="<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		aduser=trim(request("adduser"))
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
	rs("adduser")=aduser
	rs("password")=passnw
''''''''''''''
'�����û��ĵļ���
        conn.execute("update [user] set usergroupid=1,userclass='"&groupsname&"',titlepic='"&titlepic&"' where username='"&trim(request("oldusername"))&"'")	'
	body="<li>����Ա���ϸ��³ɹ������ס������Ϣ��<br> ����Ա��"&request("username2")&" <BR> ��   �룺"&request("password2")&" <a href=?>�� <font color=red>����</font> ��</a>"
	rs.update
	
	end if
	rs.close
	set rs=nothing
end sub


sub addadmin()
%> 
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա��������ӹ���Ա
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=text name="username2">  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2">  (����ע�����벻ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><input type=text name="username1">  (��ѡ����д�������޸�)
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%
end sub

sub savenew()
dim adminuserid
	if request.form("username2")="" then
	body="�������̨��¼�û�����"
	exit sub
	end if
	if request.form("username1")="" then
	body="������ǰ̨��¼�û�����"
	exit sub
	end if
	if request.form("password2")="" then
	body="�������̨��¼���룡"
	exit sub
	end if

	set rs=conn.execute("select userid from [user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
	body="��������û�������һ����Ч��ע���û���"
	exit sub
        else
        adminuserid=rs(0)
	end if

	set rs=conn.execute("select username from admin where username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
	body="��������û����Ѿ��ڹ����û��д��ڣ�"
	exit sub
	end if
	conn.execute("update [user] set usergroupid=1 , userclass='"&groupsname&"',titlepic='"&titlepic&"' where userid="&adminuserid&" ")
	conn.execute("insert into admin (username,[password],adduser) values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""))&"','"&replace(request.form("username1"),"'","")&"')")
	body="�û�ID:"&adminuserid&" ��ӳɹ������ס�¹���Ա��̨��¼��Ϣ�������޸��뷵�ع���Ա����"
end sub

sub userinfo()
dim menu(7,10)
menu(0,0)="��������"
menu(0,1)="<a href=admin_setting.asp target=main>����������Ϣ</a>"
menu(0,2)="<a href=admin_ads.asp target=main>��̳�������</a>"
menu(0,3)="<a href=bbseven.asp?action=batch target=_blank>��̳��־����</a>"
menu(0,4)="<a href=admin_inputfile.asp target=main>��ʼ��Ϣ����</a>"

menu(1,0)="��̳����"
menu(1,1)="<a href=admin_board.asp?action=add target=main>��̳�������</a> | <a href=admin_board.asp target=main>����</a>"
menu(1,2)="<a href=admin_board.asp?action=permission target=main>��̳Ȩ�޹���</a>"
menu(1,3)="<a href=admin_boardunite.asp target=main>�ϲ���̳����</a>"
menu(1,4)="<a href=admin_update.asp target=main>������̳����</a>"
menu(1,5)="<a href=z_link_post.asp target=_blank>������̳����</a>"
menu(1,6)="<a href=announcements.asp?boardid=0&action=AddAnn target=_blank>������̳���� | ����</a>"
menu(1,7)="<a href=z_admin_menpai.asp target=main>��̳���ɹ���</a>"

menu(2,0)="�û�����"
menu(2,1)="<a href=admin_user.asp target=main>�û���Ϣ����</a>"
menu(2,2)="<a href=admin_grade.asp?action=add target=main>��̳�ȼ����</a> | <a href=admin_grade.asp target=main>����</a>"
menu(2,3)="<a href=admin_wealth.asp target=main>�û���������</a>"
menu(2,4)="<a href=admin_message.asp target=main>��̳���Ź���</a>"
menu(2,5)="<a href=admin_group.asp?action=addgroup target=main>�û������</a> | <a href=admin_group.asp target=main>����</a>"
menu(2,6)="<a href=admin_admin.asp?action=add target=main>����Ա���</a> | <a href=admin_admin.asp target=main>����</a>"
menu(2,7)="<a href=admin_mailist.asp target=main>�ʼ��б�</a> | <a href=admin_mailout.asp target=main>�б���</a>"
menu(2,8)="<a href=admin_update.asp?action=updateuser target=main>�����û�����</a>"

menu(3,0)="�������������"
menu(3,1)="<a href=admin_alldel.asp target=main>����ɾ��</a>"
menu(3,2)="<a href=admin_alldel.asp?action=moveinfo target=main>�����ƶ�</a>"
menu(3,3)="<a href=recycle.asp target=_blank>����վ����</a>"
menu(3,4)="<a href=admin_postdata.asp?action=Nowused target=main>��ǰ�������ݱ����</a>"
menu(3,5)="<a href=admin_postdata.asp target=main>���ݱ������ת��</a>"

menu(4,0)="�������"
menu(4,1)="<a href=admin_color.asp target=main>��̳���CSS����</a>"
menu(4,2)="<a href=admin_pic.asp target=main>����ͼƬ����</a>"
menu(4,3)="<a href=admin_skin.asp?action=news target=main>����ģ�����</a> | <a href=admin_skin.asp target=main>����</a>"
menu(4,4)="<a href=admin_loadskin.asp target=main>CSS��񵼳�</a> | <a href=admin_loadskin.asp?action=load target=main>����</a>"

menu(5,0)="�滻/���ƴ���"
menu(5,1)="<a href=admin_badword.asp?reaction=badword target=main>���ӹ����ַ�</a>"
menu(5,2)="<a href=admin_badword.asp?reaction=splitreg target=main>ע������ַ�</a>"
menu(5,3)="<a href=admin_lockip.asp?action=add target=main>IP�����޶����</a> | <a href=admin_lockip.asp target=main>����</a>"
menu(5,4)="<a href=admin_address.asp?action=add target=main>��̳IP�����</a> | <a href=admin_address.asp target=main>����</a>"
menu(5,5)="<a href=admin_address.asp?action=upip target=main>����IP��</a>"

menu(6,0)="���ݴ���(Access)"
menu(6,1)="<a href=admin_data.asp?action=CompressData target=main>ѹ�����ݿ�</a>"
menu(6,2)="<a href=admin_data.asp?action=BackupData target=main>�������ݿ�</a>"
menu(6,3)="<a href=admin_data.asp?action=RestoreData target=main>�ָ����ݿ�</a>"
menu(6,4)="<a href=admin_data.asp?action=SpaceSize target=main>ϵͳ�ռ�ռ��</a>"

menu(7,0)="�ļ�����"
menu(7,1)="<a href=admin_upUserface.asp target=main>�ϴ�ͷ�����</a>"
menu(7,2)="<a href=admin_uploadlist.asp target=main>�ϴ��ļ�����</a>"
menu(7,3)="<a href=admin_bbsface.asp?orders=1 target=main>ע��ͷ�����</a>"
menu(7,4)="<a href=admin_bbsface.asp?orders=2 target=main>�����������</a>"
menu(7,5)="<a href=admin_bbsface.asp?orders=3 target=main>�����������</a>"
dim j,tmpmenu,menuname,menurl
set rs=conn.execute("select * from admin where id="&request("id"))
%>
<form action="admin_admin.asp?action=updat" method=post name=adminflag>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr> 
<th height=25><b>����ԱȨ�޹���</b>(��ѡ����Ӧ��Ȩ�޷��������Ա<%=rs("username")%>)
</th>
</tr>
<tr> 
<td height=25 class="forumHeaderBackgroundAlternate"><b>>>ȫ��Ȩ��</b></td></tr>
<tr><td class=forumrow>
<%for i=0 to ubound(menu,1)%>
<b><%=menu(i,0)%></b><br>
<%
on error resume next
for j=1 to ubound(menu,2)
if isempty(menu(i,j)) then exit for
tmpmenu=split(menu(i,j),",")
menuname=tmpmenu(0)
menurl=tmpmenu(1)
%>
<input type="checkbox" name="flag" <% if instr(session("flag"),"26")=0 then response.write "disabled=true" %> value="<%=i&j%>" <% if instr(rs("flag"),i&j)<>0 then response.write "checked" %>><a href="<%=menurl%>"><%=menuname%></a>&nbsp;&nbsp;
<%next%><br><br>
<%next%>
<input type=hidden name=id value="<%=request("id")%>">
<input type="submit" name="Submit" value="����">������<input name=chkall type=checkbox value=on onclick=CheckAll(this.form)>ѡ������Ȩ��
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end sub

sub update()
'00,01,02,03,10,11,12,13,14,15,16,17,20,21,22,23,24,25,26,27,28,30,31,32,33,34,35,40,41,42,43,44,50,51,52,53,54,55,60,61,62,63,64,70,71,72,73,74,75
set rs=server.createobject("adodb.recordset")
sql="select * from admin where id="&request("id")
rs.open sql,conn,1,3
if not rs.eof and not rs.bof then
rs("flag")=request("flag")
body="<li>����Ա���³ɹ������ס������Ϣ��"
rs.update
if rs("username")=membername then session("flag")=request("flag")
end if
rs.close
set rs=nothing
end sub
%>