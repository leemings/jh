<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%     
'=========================================================
' File: boardpermission.asp
' Version:5.0 Final
' Date:  2002-10-20
' Script Written by sunwin
'=========================================================
 dim orders

	stats="��̳�û���Ȩ�޲�ѯ"
	if boardid=0 then
	call nav()
	call head_var(2,0,"","")
	else
	call nav()
	call head_var(1,BoardDepth,0,0)
	end if
	if Cint(GroupSetting(39))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���������̳�û���Ȩ�޵�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
		founderr=true
	end if
	if master then founderr=false
	if founderr then
		call dvbbs_error()
	else
           if not isInteger(request("orders")) or request("orders")="" then
		orders=1
	    else
		orders=request("orders")
	   end if
		call permission()
		if founderr then call dvbbs_error()
		call activeonline()
	end if
	call footer()
	REM ��ʾ������Ϣ---Headinfo

function groupnum()
dim tmprs
	sql="Select count(*) from usergroups "
set tmprs=conn.execute(sql) 
groupnum=tmprs(0) 
set tmprs=nothing 
if isnull(groupnum) then groupnum=0
end function 

	sub permission()
dim trs,ars,Groupname
%>
<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=1 >���Ȩ��<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=2 >����Ȩ��<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=3 >���ӹ���Ȩ��<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=4 >����Ȩ��<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=5 >����Ȩ��<a></th>
<th height="25" width=16% id=tabletitlelink><a href=?boardid=<%=request("boardid")%>&orders=6 >����Ȩ��<a></th>
</tr>
</table>
<%
       select case orders
	case 1
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>���Ȩ��</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>���������̳</td>
<td height="25" width=20% class=tablebody2>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)</td>
<td height="25" width=20% class=tablebody2>���Բ鿴�����˷���������</td>
<td height="25" width=20% class=tablebody2>���������������</td>
</tr>
<%
	case 2
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=14 class=tablebody1>����Ȩ��</td>
</tr>
<tr>
<td height="25" width=22% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=6% class=tablebody2>���Է���������</td>
<td height="25" width=6% class=tablebody2>���Իظ��Լ�������</td>
<td height="25" width=6% class=tablebody2>���Իظ������˵�����</td>
<td height="25" width=6% class=tablebody2>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?</td>
<td height="25" width=6% class=tablebody2>�������������Ǯ</td>
<td height="25" width=6% class=tablebody2>�����ϴ�����</td>
<td height="25" width=6% class=tablebody2>һ������ϴ��ļ�����</td>
<td height="25" width=6% class=tablebody2>һ������ϴ��ļ�����</td>
<td height="25" width=6% class=tablebody2>�ϴ��ļ���С����</td>
<td height="25" width=6% class=tablebody2>���Է�����ͶƱ</td>
<td height="25" width=6% class=tablebody2>���Բ���ͶƱ</td>
<td height="25" width=6% class=tablebody2>���Է���С�ֱ�</td>
<td height="25" width=6% class=tablebody2>����С�ֱ������Ǯ</td>
</tr>
<%
	case 3
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>���ӹ���Ȩ��</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>���Ա༭�Լ�������</td>
<td height="25" width=20% class=tablebody2>����ɾ���Լ�������</td>
<td height="25" width=20% class=tablebody2>�����ƶ��Լ������ӵ�������̳</td>
<td height="25" width=20% class=tablebody2>���Դ�/�ر��Լ�����������</td>
</tr>
<%
	case 4
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=11 class=tablebody1>����Ȩ��</td>
</tr>
<tr>
<td height="25" width=* class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=9% class=tablebody2>����������̳</td>
<td height="25" width=9% class=tablebody2>����ʹ��'���ͱ�ҳ������'����</td>
<td height="25" width=9% class=tablebody2>�����޸ĸ�������</td>
<td height="25" width=9% class=tablebody2>�Ƿ���������ӵ�Ȩ��</td>
<td height="25" width=9% class=tablebody2>�Ƿ��н���������̳��Ȩ��</td>
<td height="25" width=9% class=tablebody2>�Ƿ�����̳�ļ�����Ȩ��</td>
<td height="25" width=9% class=tablebody2>���Խ��������̶ܹ�����</td>
<td height="25" width=9% class=tablebody2>���������̳�¼�</td>
<td height="25" width=9% class=tablebody2>���������̳չ��</td>
<td height="25" width=9% class=tablebody2>�Ƿ�������������ɫȨ��</td>
</tr>
<%
	case 5
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=18 class=tablebody1>����Ȩ��</td>
</tr>
<tr>
<td height="25" width=10% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=5% class=tablebody2>����ɾ������������</td>
<td height="25" width=5% class=tablebody2>�����ƶ�����������</td>
<td height="25" width=5% class=tablebody2>���Դ�/�ر�����������</td>
<td height="25" width=5% class=tablebody2>���Թ̶�/����̶�����</td>
<td height="25" width=5% class=tablebody2>���Խ���/�ͷ������û�</td>
<td height="25" width=5% class=tablebody2>���Խ���/�ͷ��û�</td>
<td height="25" width=5% class=tablebody2>���Ա༭����������</td>
<td height="25" width=5% class=tablebody2>���Լ���/�����������</td>
<td height="25" width=5% class=tablebody2>���Է�������</td>
<td height="25" width=5% class=tablebody2>���Թ�����</td>
<td height="25" width=5% class=tablebody2>���Թ���С�ֱ�</td>
<td height="25" width=5% class=tablebody2>��������/����/��������û�</td>
<td height="25" width=5% class=tablebody2>����ɾ���û�1��10������������</td>
<td height="25" width=5% class=tablebody2>���Բ鿴����IP����Դ</td>
<td height="25" width=5% class=tablebody2>�����޶�IP����</td>
<td height="25" width=5% class=tablebody2>���Թ����û�Ȩ��</td>
<td height="25" width=5% class=tablebody2>��������ɾ�����ӣ�ǰ̨��</td>
</tr>
<%
	case 6
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>����Ȩ��</td>
</tr>
<tr>
<td height="25" width=12.5% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=12.5% class=tablebody2>���Է��Ͷ���</td>
<td height="25" width=12.5% class=tablebody2>��෢���û�</td>
<td height="25" width=12.5% class=tablebody2>�������ݴ�С����</td>
<td height="25" width=12.5% class=tablebody2>�����С����</td>
</tr>
<%
	case else
%><table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr  align=center>
<td colspan=5 class=tablebody1>����Ȩ��</td>
</tr>
<tr>
<td height="25" width=20% class=tablebody2>�û�����(<%=groupnum()%>)</td>
<td height="25" width=20% class=tablebody2>���������̳</td>
<td height="25" width=20% class=tablebody2>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)</td>
<td height="25" width=20% class=tablebody2>���Բ鿴�����˷���������</td>
<td height="25" width=20% class=tablebody2>���������������</td>
</tr>
<%	end select

	set trs=conn.execute("select * from usergroups order by usergroupid")
        do while not trs.eof
        groupname=trs("title")

	set ars=conn.execute("select * from BoardPermission where BoardID="&boardid&" and GroupID="&trs("UserGroupID"))
          if  not ars.eof then
        GroupSetting=split(ars("PSetting"),",")
          else
         GroupSetting=split(trs("GroupSetting"),",")
         end if

       select case orders
	case 1
%>
<tr> 
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(0)=1 then%>��<%end if%><%if GroupSetting(0)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(1)=1 then%>��<%end if%><%if GroupSetting(1)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(2)=1 then%>��<%end if%><%if GroupSetting(2)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(41)=1 then%>��<%end if%><%if GroupSetting(41)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
</tr>
<%
	case 2
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(3)=1 then%>��<%end if%><%if GroupSetting(3)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(4)=1 then%>��<%end if%><%if GroupSetting(4)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(5)=1 then%>��<%end if%><%if GroupSetting(5)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(6)=1 then%>��<%end if%><%if GroupSetting(6)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(47)%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(7)=1 then%>��<%end if%><%if GroupSetting(7)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%><%if GroupSetting(7)=2 then%><font color=<%=Forum_body(8)%>>��</font><%end if%><%if GroupSetting(7)=3 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(40)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(50)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(44)%></B> KB</td>
<td height="25"  class=tablebody1><B><%if GroupSetting(8)=1 then%>��<%end if%><%if GroupSetting(8)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(9)=1 then%>��<%end if%><%if GroupSetting(9)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(17)=1 then%>��<%end if%><%if GroupSetting(17)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(46)%></B></td>
</tr>
<%
	case 3
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(10)=1 then%>��<%end if%><%if GroupSetting(10)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(11)=1 then%>��<%end if%><%if GroupSetting(11)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(12)=1 then%>��<%end if%><%if GroupSetting(12)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(13)=1 then%>��<%end if%><%if GroupSetting(13)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
</tr>
<%
	case 4
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(14)=1 then%>��<%end if%><%if GroupSetting(14)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(15)=1 then%>��<%end if%><%if GroupSetting(15)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(16)=1 then%>��<%end if%><%if GroupSetting(16)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(36)=1 then%>��<%end if%><%if GroupSetting(36)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(37)=1 then%>��<%end if%><%if GroupSetting(37)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(48)=1 then%>��<%end if%><%if GroupSetting(48)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(38)=1 then%>��<%end if%><%if GroupSetting(38)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(39)=1 then%>��<%end if%><%if GroupSetting(39)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(49)=1 then%>��<%end if%><%if GroupSetting(49)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(51)=1 then%>��<%end if%><%if GroupSetting(51)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
</tr>
<%
	case 5
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(18)=1 then%>��<%end if%><%if GroupSetting(18)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(19)=1 then%>��<%end if%><%if GroupSetting(19)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(20)=1 then%>��<%end if%><%if GroupSetting(20)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(21)=1 then%>��<%end if%><%if GroupSetting(21)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(22)=1 then%>��<%end if%><%if GroupSetting(22)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(43)=1 then%>��<%end if%><%if GroupSetting(43)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(23)=1 then%>��<%end if%><%if GroupSetting(23)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(24)=1 then%>��<%end if%><%if GroupSetting(24)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(25)=1 then%>��<%end if%><%if GroupSetting(25)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(26)=1 then%>��<%end if%><%if GroupSetting(26)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(27)=1 then%>��<%end if%><%if GroupSetting(27)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(28)=1 then%>��<%end if%><%if GroupSetting(28)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(29)=1 then%>��<%end if%><%if GroupSetting(29)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(30)=1 then%>��<%end if%><%if GroupSetting(30)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(31)=1 then%>��<%end if%><%if GroupSetting(31)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(42)=1 then%>��<%end if%><%if GroupSetting(42)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(45)=1 then%>��<%end if%><%if GroupSetting(45)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
</tr>
<%
	case 6
%>
<tr>
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(32)=1 then%>��<%end if%><%if GroupSetting(32)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(33)%></B></td>
<td height="25"  class=tablebody1><B><%=GroupSetting(34)%></B> byte</td>
<td height="25"  class=tablebody1><B><%=GroupSetting(35)%></B> KB</td>
</tr>
<%
	case else
%>
<tr> 
<td height="25"  class=tablebody1><%=groupname%></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(0)=1 then%>��<%end if%><%if GroupSetting(0)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(1)=1 then%>��<%end if%><%if GroupSetting(1)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(2)=1 then%>��<%end if%><%if GroupSetting(2)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
<td height="25"  class=tablebody1><B><%if GroupSetting(41)=1 then%>��<%end if%><%if GroupSetting(41)=0 then%><font color=<%=Forum_body(8)%>>��</font><%end if%></B></td>
</tr>
<%
	end select
trs.movenext
loop
set trs=nothing
set ars=nothing
%>
</table>
<% end sub %>