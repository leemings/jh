<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!--#include file="md5.asp"-->
<%
'=========================================================
' File: smallpaper.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim cansmallpaper
cansmallpaper=false
stats="����С�ֱ�"
if boardid=0 then
	founderr=true
	Errmsg=Errmsg+"<br><li>��ѡ����Ҫ����С�ֱ��İ��棡"
end if

if Cint(GroupSetting(17))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�з���С�ֱ���Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
else
	if not founduser then
		membername="����"
	end if
	cansmallpaper=true
end if

if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
		end if
	end if
end if
'###################��������޸Ŀ�ʼ(asilas����)##################
if cint(Board_Setting(51))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkviplogin(membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>VIP��Աר�ð���</font>����ȷ�����������Ƿ���ϡ�"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(52))<>0 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ������̳����<a href=login.asp>��¼</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		dim sexshow
		if cint(Board_Setting(52))=1 then
		sexshow="Ů��"
		elseif cint(Board_Setting(52))=2 then
		sexshow="����"
		end if
		if chksexlogin(cint(Board_Setting(52)),membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳����Ϊ<font color=red>"&sexshow&"��̳����</font>����ȷ�������Ա��Ƿ���ϡ�"
		founderr=true
		end if
	end if
end if
'####################��������޸Ľ���(asilas����)#################

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	if request("action")="savepaper" then
		call savepaper()
	else
		call main()
	end if
	call activeonline()
	if founderr then call dvbbs_error()
end if
call footer()
sub main()
conn.execute("delete from smallpaper where datediff('d',s_addtime,Now())>1")
%>
<form action="smallpaper.asp?action=savepaper" method="post"> 
        <table cellpadding=6 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2>
    ����ϸ��д������Ϣ</th></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�û���</b></td>
    <td class=tablebody1 valign=middle><INPUT name=username type=text value="<%=membername%>"> &nbsp; <a href="reg.asp">û��ע�᣿</a></td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�� ��</b></font></td>
    <td class=tablebody1 valign=middle><INPUT name=password type=password value="<%=memberword%>"> &nbsp; <a href="lostpass.asp">�������룿</a></td></tr>
    <tr>
    <td class=tablebody1 valign=middle><b>�� ��</b>(���80��)</td>
    <td class=tablebody1 valign=middle><INPUT name="title" type=text size=60></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30%>
<b>�� ��</b><BR>
�ڱ��淢��С�ֱ���������<font color="<%=Forum_body(8)%>"><b><%=GroupSetting(46)%></b></font>Ԫ����<br>
<font color="<%=Forum_body(8)%>"><b>48</b></font>Сʱ�ڷ����С�ֱ��������ȡ<font color="<%=Forum_body(8)%>"><b>5</b></font>��������ʾ����̳��<br>
<li>HTML��ǩ�� <%if Forum_Setting(35)=0 then%>������<%else%>����<%end if%>
<li>UBB ��ǩ�� <%if Forum_Setting(34)=0 then%>������<%else%>����<%end if%>
<li>���ݲ��ó���500��
</td>
    <td class=tablebody1 valign=middle>
<textarea class="smallarea" cols="60" name="Content" rows="8" wrap="VIRTUAL"></textarea>
<INPUT name="boardid" type=hidden value="<%=boardid%>">
                </td></tr>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name="submit" value="�� ��"></td></tr></table>
</form>
<%end sub%>
<%
sub savepaper()
dim username
dim password
dim title
dim content
userName=Checkstr(trim(request.form("username")))
PassWord=Checkstr(trim(request.form("password")))
title=Checkstr(trim(request.form("title")))
Content=Checkstr(request.form("Content"))
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
	FoundErr=True
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>����������"
	FoundErr=True
end if
if title="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>���ⲻӦΪ�ա�"
elseif strLength(title)>80 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>���ⳤ�Ȳ��ܳ���80"
end if
if content="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>û����д���ݡ�"
	FoundErr=true
elseif strLength(content)>500 then
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���500"
	FoundErr=true
end if

'���˲���������֤�û�
if not founderr and cansmallpaper then
	if PassWord<>memberword then
		password=md5(password)
	end if
	set rs=server.createobject("adodb.recordset")
	sql="Select userWealth From [User] Where UserName='"&UserName&"' and UserPassWord='"&PassWord&"'"
	rs.open sql,conn,1,3
	if not(rs.eof and rs.bof) then
		if Clng(rs("UserWealth"))<Clng(GroupSetting(46)) then
   			ErrMsg=ErrMsg+"<Br>"+"<li>��û���㹻�Ľ�Ǯ������С�ֱ����쵽��̳����ˮ�ɡ�"
   			FoundErr=true
		else
			rs("UserWealth")=rs("UserWealth")-Cint(GroupSetting(46))
			rs.update
		end if
	else
		errmsg=errmsg+"<br>"+"<li>�����û����������ڣ���������������󣬻��������ʺ��ѱ�����Ա������"
		founderr=true
	end if
	rs.close
	set rs=nothing
end if
if founderr then
	exit sub
else
	sql="insert into smallpaper (s_boardid,s_username,s_title,s_content) values ("&boardid&",'"&username&"','"&title&"','"&content&"')"
	conn.execute(sql)
	sucmsg="<li>���ɹ��ķ�����С�ֱ���"
	call dvbbs_suc()
end if
end sub
%>