<%
Rem ==========��̳��ҳ����=========
Rem �ж��û�ϵͳ��Ϣ
function usersysinfo(info,getinfo)
dim usersys
dim userlang
if instr(1,info,");")>0 then
	userlang=right(info,len(info)-instr(1,info,");")-1)
else
	userlang="δ֪"
end if
if instr(info,";")>0 then
	usersys=split(info,";")
	if ubound(usersys)>=2 then
		usersys(1)=replace(usersys(1),"MSIE","Internet Explorer")
		usersys(2)=replace(usersys(2),")","")
		usersys(2)=replace(usersys(2),"NT 5.1","XP")
		usersys(2)=replace(usersys(2),"NT 5.0","2000")
		usersys(2)=replace(usersys(2),"9x","Me")
		usersys(1)="� �� ����" & Trim(usersys(1))
		usersys(2)="����ϵͳ��" & Trim(usersys(2))
		if getinfo=1 then
			usersysinfo=usersys(1)
		elseif getinfo=2 then
			usersysinfo=usersys(2)
		else
			usersysinfo=userlang
		end if
	else
		if getinfo=1 then
			usersysinfo="� �� ����δ֪"
		elseif getinfo=2 then
			usersysinfo="����ϵͳ��δ֪"
		else
			usersysinfo=userlang
		end if
	end if
else
	if getinfo=1 then
		usersysinfo="δ֪"
	elseif getinfo=2 then
		usersysinfo="δ֪"
	else
		usersysinfo=userlang
	end if
end if
end function

function online(boardid,inclChild)
if boardid=0 then
	sql="Select count(*) from online where userid>0"
else
	if not inclChild then
		sql="Select count(*) from online where userid>0 and boardid="&boardid
	else
		sql="Select count(*) from online where userid>0 and (boardid="&boardid&" or boardid in (select boardid from board where parentStr='"&boardid&"' or parentstr like '"&boardid&",%' or parentstr like '%,"&boardid&",%' or parentstr like '%,"&boardid&"'))"
	end if
end if
on error resume next
online=conn.execute(sql)(0)
err.clear
on error goto 0
if isnull(online) then online=0
end function 

function guest(boardid,inclChild)
if boardid=0 then
	sql="Select count(*) from online where userid=0"
else
	if not inclChild then
		sql="Select count(*) from online where userid=0 and boardid="&boardid
	else
		sql="Select count(*) from online where userid=0 and (boardid="&boardid&" or boardid in (select boardid from board where parentStr='"&boardid&"' or parentstr like '"&boardid&",%' or parentstr like '%,"&boardid&",%' or parentstr like '%,"&boardid&"'))"
	end if
end if
on error resume next
guest=conn.execute(sql)(0)
err.clear
on error goto 0
if isnull(guest) then guest=0
end function

sub onlineuser(online_u,online_g,boardid)
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
response.write "<tr><td colspan=2 class=tablebody1><table cellpadding=2 cellspacing=1 border=0 width=""100%"" style=""word-break:break-all;""><tr>"
end if
dim online_face
dim sip,acip,userip
dim NowStats,ActiveTime,Binfo,UComeFrom,BrStr,OnlineTime
dim SexColor

if cint(online_u)=1 then
	i=0
	'�û���Ϣ
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,browser,usersex,uservip from online where userid>0 order by lastimebk desc"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,browser,usersex,uservip from online where boardid="&boardid&" and userid>0 order by lastimebk desc"
	end if
	set rs=conn.execute(sql)
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	'if membername=rs(0) then
		'Brstr="&#13;&#10;"
	'else
		Brstr="<br>"
	'end if
	if Cint(forum_setting(33))=1 then
		NowStats="Ŀǰλ�ã�" & htmlencode(rs(2)) & Brstr
	else
		NowStats=""
	end if
	if Cint(forum_setting(34))=1 then
		ActiveTime="����ʱ�䣺" & rs(6) & Brstr & "�ʱ�䣺" & rs(7) & Brstr
	else
		ActiveTime=""
	end if
	OnlineTime="����ʱ�䣺" & int(datediff("s",rs(6),now())/60) & " ����" & Brstr
	if Cint(forum_setting(35))=1 then
		Binfo=usersysinfo(rs(9),2) & "," & usersysinfo(rs(9),3) & Brstr & usersysinfo(rs(9),1) & Brstr
	else
		Binfo=""
	end if
	if Cint(GroupSetting(30))=0 then
		userip="��ʵ�ɣУ������ñ���"
	else
		if acip <> "" then
			userip="��ʵ�ɣУ�" & acip
		else
			userip="��ʵ�ɣУ�" & sip
		end if
	end if
	SexColor=iif(rs(10)," class=cman "," class=cgirl ")
 	if cint(forum_setting(36))=1 then
		if Cint(GroupSetting(30))=0 then
			UComeFrom=Brstr & "�û���Դ�������ñ���"
		else
			if acip<>"" then
				UComeFrom=Brstr & "�û���Դ��" & address(acip)
			else
				UComeFrom=Brstr & "�û���Դ��" & address(sip)
			end if
		end if
	else
		UComeFrom=""
	end if
	online_face="<a href=javascript:openScript('messanger.asp?action=new&touser="&rs(0)&"',500,400)>"
	select case rs(3)
	case 1
	online_face=online_face&"<img src="&Forum_info(7)&Forum_pic(0)&" alt=""��"&rs(0)&"����Ϣ"" height=18 border=0></a>"
	case 2
	online_face=online_face&"<img src="&Forum_info(7)&Forum_pic(16)&" alt=""��"&rs(0)&"����Ϣ"" height=18 border=0></a>"
	case 3
	online_face=online_face&"<img src="&Forum_info(7)&Forum_pic(1)&" alt=""��"&rs(0)&"����Ϣ"" height=18 border=0></a>"
	case 8
	online_face=online_face&"<img src="&Forum_info(7)&Forum_pic(2)&" alt=""��"&rs(0)&"����Ϣ"" height=18 border=0></a>"
	case else
	online_face=online_face&"<img src="&Forum_info(7)&Forum_pic(3)&" alt=""��"&rs(0)&"����Ϣ"" height=18 border=0></a>"
	end select

	if membername=rs(0) then
		response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank><font color="&Forum_pic(5)&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom &""">"&htmlencode(rs(0))&"</font></a>"
		if rs("uservip")=1 then response.write "<font color="&forum_body(8)&" style=""font-size:6pt;font-family:Arial Black""><B>VIP</b></font>"
		response.write "</td>"
	else
		if rs(4)=1 then
			if superboardmaster or master then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬������"">"&htmlencode(rs(0))&"</a>"
				if rs("uservip")=1 then response.write "<font color="&forum_body(8)&" style=""font-size:6pt;font-family:Arial Black""><b>VIP</b></font>"
				response.write "</td>"
			else
				response.write "<td width=""14%""><img src="&Forum_info(7)&Forum_pic(4)&" height=18>&nbsp;<a href=# target=_blank title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom &""">�����Ա</a></td>"
			end if
		else
			if rs(4)=3 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬��æµ"">"&htmlencode(rs(0))&"</a>"
			elseif rs(4)=4 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬�����ϻ���"">"&htmlencode(rs(0))&"</a>"
			elseif rs(4)=5 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬���뿪"">"&htmlencode(rs(0))&"</a>"
			elseif rs(4)=6 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬�������绰"">"&htmlencode(rs(0))&"</a>"
			elseif rs(4)=7 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬������Ͳ�"">"&htmlencode(rs(0))&"</a>"
			elseif rs(4)=8 then
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom  & Brstr & "��ǰ״̬������"">"&htmlencode(rs(0))&"</a>"
			else
				response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank "&SexColor&" title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom &""">"&htmlencode(rs(0))&"</a>"
			end if
			if rs("uservip")=1 then response.write "<font color="&forum_body(8)&" style=""font-size:6pt;font-family:Arial Black""><b>VIP</b></font>"
			response.write "</td>"
		end if
	end if
	if i=5 then response.write "</tr><tr>"
	if i>5 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if cint(online_g)=1 then
	dim onlineusername
	i=0
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id,browser,usersex,uservip from online where userid=0 order by lastimebk desc"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id,browser,usersex,uservip from online where boardid="&boardid&" and userid=0 order by lastimebk desc"
	end if
	set rs=conn.execute(sql)
  
	online_face="<img src="&Forum_info(7)&Forum_pic(4)&" height=18>"
	if not (rs.eof and rs.eof) then
		response.write "</tr><tr>"
	end if
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	if trim(session("userid"))<>"" and isnumeric(session("userid")) then
		if int(session("userid"))=int(rs(9)) then
			Brstr="&#13;&#10;"
			onlineusername="<font color="&Forum_pic(5)&">����</font>"
		else
			Brstr="<br>"
			onlineusername="����"
		end if
	else
		Brstr="<br>"
		onlineusername="����"
	end if
	if Cint(forum_setting(33))=1 then
		NowStats="Ŀǰλ�ã�" & htmlencode(rs(2)) & Brstr
	else
		NowStats=""
	end if
	if Cint(forum_setting(34))=1 then
		ActiveTime="����ʱ�䣺" & rs(6) & Brstr & "�ʱ�䣺" & rs(7) & Brstr
	else
		ActiveTime=""
	end if
	OnlineTime="����ʱ�䣺" & int(datediff("s",rs(6),now())/60) & " ����" & Brstr
	if Cint(forum_setting(35))=1 then
		Binfo=usersysinfo(rs(10),2) & "," & usersysinfo(rs(9),3) & Brstr & usersysinfo(rs(10),1) & Brstr
	else
		Binfo=""
	end if
	if Cint(GroupSetting(30))=0 then
		userip="��ʵ�ɣУ������ñ���"
	else
		if acip <> "" then
			userip="��ʵ�ɣУ�" & acip
		else
			userip="��ʵ�ɣУ�" & sip
		end if
	end if
	if cint(forum_setting(36))=1 then
		if Cint(GroupSetting(30))=0 then
			UComeFrom=Brstr & "�û���Դ�������ñ���"
		else
			if acip<>"" then
				UComeFrom=Brstr & "�û���Դ��" & address(acip)
			else
				UComeFrom=Brstr & "�û���Դ��" & address(sip)
			end if
		end if
	else
		UComeFrom=""
	end if
	response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=# target=_blank title="""& NowStats & ActiveTime & OnlineTime & Binfo & UserIP & UComeFrom &""">"&onlineusername&"</a></td>"

	if i=5 then response.write "</tr><tr>"
	if i>5 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
	response.write "</tr></TABLE>"
end if
set rs=nothing
end sub
%>