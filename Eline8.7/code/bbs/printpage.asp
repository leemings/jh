<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
' File: printpage.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim urs,usql
dim rsboard,boardsql
dim announceid
dim username
dim rootid
dim topic
dim abgcolor
dim postbuyuser
dim replyid
abgcolor="#FFFFFF"
Dim TotalUseTable
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	AnnounceID=request("id")
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

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
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

if foundErr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call announceinfo()
	if founderr then call dvbbs_error()
end if
call footer()

sub announceinfo()
    set rs=conn.execute("select title,PostTable,LockTopic,PostUserName from topic where topicID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		if rs("locktopic")=2 then
			ErrMsg=ErrMsg+"<br>"+"<li>�������Ѿ�������Աɾ����</li>"
			founderr=true
			exit sub
		end if
		TotalUseTable=rs(1)
		topic=rs(0)
		if lcase(trim(rs("PostUserName")))<>lcase(trim(membername)) then
			if Cint(GroupSetting(2))=0 then
				Errmsg=Errmsg+"<br>"+"<li>��û������ڱ���̳�鿴�����˷��������ӵ�Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
				founderr=true
				exit sub
			end if
		end if
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>��ָ�������Ӳ�����</li>"
		exit sub
	end if
	rs.close
%>
<HTML><HEAD><TITLE><%=Forum_info(0)%>--��ʾ����</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<!--#include file="inc/Forum_css.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY bgcolor="#ffffff" alink="#333333" vlink="#333333" link="#333333" topmargin="10" leftmargin="10">
<TABLE border=0 width="<%=Forum_body(12)%>" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
<b>���ı���ʽ�鿴����</b><br><br>
-&nbsp;&nbsp;<b><%=Forum_info(0)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>index.asp)<br>
--&nbsp;&nbsp;<b><%=boardtype%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>list.asp?boardid=<%=boardid%>)<br>
----&nbsp;&nbsp;<b><%=htmlencode(topic)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>dispbbs.asp?boardid=<%=boardid%>&rootid=<%=rootid%>&id=<%=announceid%>)
      </TD></TR></TBODY></TABLE> 
<br>
<hr>
<%
	Rs.open "Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID,b.postbuyuser from "&TotalUseTable&" b inner join [user] u on b.PostUserID=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" order by b.announceid",conn,1,1
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>û���ҵ�������ӡ�"
		exit sub
	end if
	do while not rs.eof
	postbuyuser=rs("postbuyuser")
	username=rs(0)
%>
<TABLE border=0 width="750" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
--&nbsp;&nbsp;���ߣ�<%=rs("username")%><br>
--&nbsp;&nbsp;����ʱ�䣺<%=rs("dateandtime")%><br><br>
--&nbsp;&nbsp;<%=htmlencode(rs("topic"))%><br>
<%=dvbcode(rs("body"),rs("UserGroupID"),1)%>
	<hr>
    </TD></TR></TBODY></TABLE> 
<%
          rs.movenext
        loop	
	rs.close
	set rs=nothing
	call activeonline()
end sub    
%>