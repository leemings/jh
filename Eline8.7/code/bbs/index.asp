<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<%'ȥ�����
'<!--#include file="z_dgconn.asp"-->
%>
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="z_Visual_const.asp"-->
<%
response.buffer=true
'=========================================================
' File: index.asp
' Version:5.0
' Date: 2002-9-21
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
Dim OnlineNum,GuestNum
Dim reBoard_Setting,BoardCount,k,ColSpanNum
Dim Board_Count
Dim TdTableWidth
OnlineNum=online(0,false)
GuestNum=guest(0,false)
Stats="��̳��ҳ"
call activeonline()
Rem ��ҳ������Ϣ
call nav()%>
<%if isnull(lastlogin) or lastlogin="" then lastlogin=now()
if Clng(Forum_OnlineNum)>Clng(Maxonline) then
	conn.execute("update config set Maxonline="&allonline()&",MaxonlineDate=Now()")
end if%>
<BR>
<table cellspacing=0 cellpadding=0 align=center width=<%=forum_body(12)%> border=0>
	<tr>
		<td><!--#include file="z_announce.asp"--></td>
	</tr>
</table>

<!--#include file="z_IndexHead.asp"-->
<%if cint(forum_setting(75))=1 then
	if cint(request("setup"))<>1 then
		if Cint(Forum_setting(4))=1 then
			call main_2()
		else
			call showmylist()
			if request.Cookies("bbslist")("Isshow")<>"0" then 
				call main_1()
			end if
		end if
	else
		call showmylist()
	end if
else
	if Cint(Forum_setting(4))=1 then
		call main_2()
	else
		call main_1()
	end if
end if

'ȥ�����
'if cint(forum_setting(12))=1 then%>
<%'	<!--#include file="z_dgshow.asp"-->%>
<%'end if
if cint(forum_setting(16))=1 then
	response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><TR height=25><Th align=left id=tabletitlelink colSpan=4 style='font-weight:normal'><b>��������</b>&nbsp;&nbsp;&nbsp;&nbsp;"
	if request("mingxing")="off" then
		response.write "[<a href=?action="&request("action")&"&mingxing=on>��ʾ��ϸ�б�</a>]</th></tr>"
	else
		response.write "[<a href=?action="&request("action")&"&mingxing=off>�ر���ϸ�б�</a>]</th></tr>"
	end if
	if request("mingxing")<>"off" then
		call mingxing()
	end if
	response.write "<tr><td colSpan=4 class=tablebody2>"&mingxingtop(5)&"</td></tr></table>"
	response.write "<BR>"
end if
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"
response.write "<TR><Th colSpan=2 align=left height=25>������̳  | <a href=""z_link_apply.asp""><font color=#ffffff>��վ��Ϣ</font></a> | <a href=""z_link_post.asp""><font color=#ffffff>��������</font></a>&nbsp;&nbsp;&nbsp;"
if cint(forum_setting(77))>1 then
	response.write "<SPAN style=""FONT-SIZE: 9px; FONT-FAMILY: webdings""><script>var stopflag=0;</script>"
	response.write "<FONT onmouseup=""document.all.lmforum.scrollAmount=6;"" onmousedown=""document.all.lmforum.direction='left';document.all.lmforum.scrollAmount=10;if(stopflag==1){stopflag=0;document.all.lmforum.start();}"" onmouseover=""if(document.all.lmforum.direction == 'left')document.all.lmforum.scrollAmount=6;"" style=""CURSOR: hand"" onmouseout=document.all.lmforum.scrollAmount=2;>7</FONT>&nbsp;"
	response.write "<FONT style=""CURSOR: default"" onclick=""if(stopflag==1){stopflag=0;document.all.lmforum.start();document.all.startbutton.style.cursor='default';document.all.stopbutton.style.cursor='hand';}"" id=""startbutton"">4</FONT>&nbsp;"
	response.write "<FONT style=""CURSOR: hand"" onclick=""if(stopflag==0){stopflag=1;document.all.lmforum.stop();document.all.startbutton.style.cursor='hand';document.all.stopbutton.style.cursor='default';}"" id=""stopbutton"">&lt;</FONT>&nbsp;"
	response.write "<FONT onmouseup=document.all.lmforum.scrollAmount=6; onmousedown=""document.all.lmforum.direction='right';document.all.lmforum.scrollAmount=10;if(stopflag==1){stopflag=0;document.all.lmforum.start();}"" onmouseover=""if(document.all.lmforum.direction == 'right')document.all.lmforum.scrollAmount=6;"" style=""CURSOR: hand"" onmouseout=document.all.lmforum.scrollAmount=2;>8</FONT>"
	response.write "</span>"
end if
response.write "</Th></tr>"
response.write "<TR><TD vAlign=top class=tablebody1 width=100% ><table width=100% >"
myCache.name="bbslink"
if myCache.valid then
	response.write myCache.value
else
	call cache_link(forum_setting(77))
	response.write myCache.value
end if
response.write "</table></TD></TR></table>"
response.write "<br>"

'=========������ʾ��ʼ================
if cint(forum_setting(29))=1 then
	myCache.name="userbirthday"
	if myCache.valid then
		response.write myCache.value
		'myCache.MakeEmpty
	else
		dim age
		dim birthuser
		dim foundbirth
		dim showbirthday
		dim birthNum,birthday
		foundbirth=false
		birthNum=0
		sql="select birthuser from config where active=1"
		set rs=conn.execute(sql)
		if not isnull(rs(0)) or rs(0)<>"" then
			birthuser=split(rs(0),"$")
			if ubound(birthuser)<3 then
				foundbirth=false
			elseif cint(birthuser(1))<>0 then
				foundbirth=true
			elseif datediff("d",birthuser(2),Now())>0 then
				foundbirth=false
			else
				foundbirth=true
			end if
		else
			foundbirth=false
		end if
		if not foundbirth then
			set rs=conn.execute("select username,birthday from [user] where birthday<>''  order by userid") 
			do while not rs.eof
				if  isdate(rs(1)) then
					if month(rs(1))=month(Now()) and  day(rs(1))=day(Now()) then
						age=datediff("yyyy",rs(1),Now())
						birthday=birthday & "<a href=dispuser.asp?name="&rs(0)&" title=ף"&age&"�����տ��֣� target=_blank>��ף "&rs(0)&" ���տ���<img src=pic/birthday00.gif align=absmiddle border=0>��  </a> ,"
						birthNum=birthNum+1
					end if
				end if
				rs.movenext
			loop
			rs.close
			set rs=nothing
			if birthday="" then  birthday="����û�����ѹ�����..."
			if right(birthday,2)=" ," then birthday=left(birthday,len(birthday)-2)
			conn.execute("update config set birthuser='" & birthday & "$" & birthNum & "$" & Now() & "' where active=1")
		else
			birthday=birthuser(0)
			birthNum=birthuser(1)
		end if
		showbirthday="<table cellpadding=3 cellspacing=1 align=center class=tableborder1><TR ><Th align=left height=25>-=> ��������յ��û����� "&birthNum&" �ˣ�</Th></TR><TR><TD width=100% vAlign=middle class=tablebody1  >"&birthday&"</TD></TR></table><br>"
		myCache.add showbirthday,date+1
	end if
end if
'=========������ʾ����================

'=========��������Ա��ʾ��ʼ================
if cint(forum_setting(43))=1 then
	myCache.name="showadduser"
	if myCache.valid then
		response.write myCache.value
		'myCache.MakeEmpty
	else
		dim ismember,showuser,adduser,SexColor
		ismember=0
		set rs=conn.execute("select userid,username,addDate,sex from [user] where month(addDate)=month(now()) and day(addDate)=day(now()) and year(addDate)<year(now()) order by userid") 
		do while not rs.eof
			ismember=ismember+1
			SexColor=iif(rs(3)," class=cman "," class=cgirl ")
			showuser=showuser & "<a href=dispuser.asp?id="&rs(0)&" title=""��Ա ["&rs(1)&"] ������̳"&DateDiff("yyyy", rs(2), now())&"��������գ�ף�����֣�"" target=_blank "&SexColor&">�� "&rs(1)&" ��</a> ,"
			rs.movenext
		loop
		rs.close
		set rs=nothing
		if showuser="" then showuser="����û�������������..."
		if right(showuser,2)=" ," then showuser=left(showuser,len(showuser)-2)
		adduser="<table cellpadding=3 cellspacing=1 align=center class=tableborder1><TR><Th align=left height=25>-=> ������������Ա���� "&ismember&" �ˣ�</Th></TR><TR><TD width=100% vAlign=middle class=tablebody1>"&showuser&"</TD></TR></table><BR>"
		response.write adduser
		myCache.add adduser,date+1
	end if
end if
'=========��������Ա��ʾ����================

response.write "<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style=""word-break:break-all;"">"
response.write "<TR><Th align=left colSpan=2 height=25>�û�������Ϣ</Th></TR>"
response.write "<TR><TD vAlign=center class=tablebody1 height=25 width=100% >"
Dim userip,userip2
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
userip2 = Request.ServerVariables("REMOTE_ADDR")
if userip = ""  then
	response.write "������ʵ�ɣ� �ǣ�"&userip2&"��"
else
	response.write "������ʵ�ɣ� �ǣ�"&userip&"��"
end if
response.write usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),2)&"��"&usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),1)&"��"&Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
response.write "</TD><TD class=tablebody1 vAlign=top height=25 align=right ><a href=http://www.happyjh.com title=վ������� ���ѵĿռ䣡><img border=0 src=Images/img_marry/_vti_cnf/link.gif width=88 height=31></a></TD></TR></table><BR>"
if cint(forum_setting(13))=1 then
	response.write "<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style=""word-break:break-all;"">"
	response.write "<TR><Th align=left colSpan=2 height=25>-=> ��̳����վ</Th></TR>"
	response.write "<TR><TD vAlign=top class=tablebody1 height=25 width=100% >"
	response.write "<img src="&Forum_Info(7)&"recycled.gif width=13 height=16 align=absbottom>&nbsp;��ǰ��̳����վ���� <b><font color="&forum_body(8)&">"
	dim RecycleNum
	RecycleNum=conn.execute("select count(*) from topic where locktopic=2")(0)
	for i=0 to ubound(allposttable)
		RecycleNum=RecycleNum+conn.execute("select count(*) from "&allposttable(i)&" where locktopic=2 and not parentid=0")(0)
	next
	response.write RecycleNum&"</font></b> ƪ��ɾ���ӣ�����������ظ�������<a href=recycle.asp><font color="&forum_body(8)&">�����˴��鿴</font></a>��һ�����������飬������ա�"
	response.write "</TD></TR></table><BR>"
end if
response.write "<table cellpadding=5 cellspacing=1 class=tableborder1 align=center style=""word-break:break-all;""><TR><Th colSpan=2 align=left id=tabletitlelink height=25 style=""font-weight:normal""><b>��̳����ͳ��</b>&nbsp;"
if request("action")="show" then
	response.write "[<a href=?action=off>�ر���ϸ�б�</a>]"
else
	if cint(Forum_Setting(14))=1 and request("action")<>"off" then
		response.write "[<a href=?action=off>�ر���ϸ�б�</a>]"
	else
		response.write "[<a href=?action=show>��ʾ��ϸ�б�</a>]"
	end if
end if
response.write "[<a href=boardstat.asp?reaction=online>�鿴�����û�λ��</a>]</Th></TR>"
response.write "<TR>"
response.write "<TD width=100% vAlign=top class=tablebody1>  Ŀǰ��̳���ܹ��� <b>"&Forum_OnlineNum&"</b> �����ߣ�����ע���Ա <b>"&OnlineNum&"</b> �ˣ��ÿ� <b>"&GuestNum&"</b> �ˡ�<br>"
response.write "��ʷ������߼�¼�� <b>"&Maxonline&"</b> ��ͬʱ���ߣ�����ʱ���ǣ�"&formatdatetime(MaxonlineDate,1)&"&nbsp;"&formatdatetime(MaxonlineDate,4)
response.write "<BR><font color="&Forum_body(8)&">����ͼ��</font>��"
response.write "<img src="&Forum_info(7)&Forum_pic(0)&"> ����Ա   ����<img src="&Forum_info(7)&Forum_pic(16)&"> ��������   ��   <img src="&Forum_info(7)&Forum_pic(1)&"> ����   ��   <img src="&Forum_info(7)&Forum_pic(2)&"> ��̳���   ��   <img src="&Forum_info(7)&Forum_pic(3)&"> ��ͨ��Ա   ��   <img src="&Forum_info(7)&Forum_pic(4)&"> ���˻������Ա"
response.write "<br><font color="&Forum_body(8)&">�Ա���ɫ</font>�� <a class=cman>����</a> �� <a class=cgirl>Ů��</a> �� <font color="&Forum_pic(5)&">�Լ�</font> �� ���˻������Ա "
response.write "<hr size=1 color="&forum_body(27)&">"
response.write "<table width=100% border=0 cellspacing=0 cellpadding=0>"
if request("action")="off" then
	call onlineuser(0,0,0)
elseif request("action")="show" then
	call onlineuser(1,1,0)
else
	call onlineuser(Forum_Setting(14),Forum_Setting(15),0)
end if
response.write "</table></TD></TR>"
response.write "</TABLE><BR>"%>
<table cellspacing=1 cellpadding=3 width="<%=Forum_body(12)%>" border=0 align=center>
	<tr>
		<td align=center><img src="<%=Forum_info(7)&Forum_pic(6)%>" width=32 height=32 align="absmiddle">&nbsp;û���µ�����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(7)%>" width=32 height=32 align="absmiddle">&nbsp;���µ�����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(14)%>" width=32 height=32 align="absmiddle">&nbsp;û��������������̳&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(15)%>" width=32 height=32 align="absmiddle">&nbsp;��������������̳</td>
	</tr>
</table>
<%if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
	if Forum_ads(2)=1 then
		call admove()
	end if
	if Forum_ads(13)=1 then
		call fixup()
	end if
end if
call footer()


%>

<%

sub main_1()
	Dim cBoardID
	Response.Cookies("BoardList").expires= date+7
	cBoardID= Request("cBoardid")
	if isnumeric(cBoardID) then
		if request("CatLog")="N" then
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
		else
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
		end if
	end if
	dim simdisp
	simdisp=false
	k=0
	'ֻ�г����С��2�İ���
	sql="select * from board where depth<="&forum_setting(5)&" order by rootid,orders"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if not (rs.eof and rs.bof) then BoardCount=rs.recordcount else boardcount=0 end if
	do while not rs.eof
		k=k+1
		reBoard_Setting=split(rs("Board_setting"),",")
		if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
			master_1=split(rs("BoardMaster"), "|")
			for i = 0 to ubound(master_1)
				if i>6 then
					master_2=master_2
				else
					master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
				end if
			next
			if i>7 then master_2=master_2 & "<font color=gray>More...</font>"
		else
			master_2="����"
		end if
		if rs("ParentID")=0 then
			if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
				ColSpanNum=reBoard_Setting(41)
				if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
				simdisp=true
			elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
				simdisp=false
				ColSpanNum=2
			else
				if Cint(reBoard_Setting(39))=1 then
					ColSpanNum=reBoard_Setting(41)
					if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
					simdisp=true
				else
					simdisp=false
					ColSpanNum=2
				end if
			end if
			response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
			response.write "<TR><Th colSpan="&ColSpanNum&" height=25 align=left id=TableTitleLink>&nbsp;"
			if simdisp then
				response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=Y"" title=""չ����̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>"
			else
				response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=N"" title=""�ر���̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>"
			end if
			response.write "</Th></TR>"
			Board_Count=0
		else
			if simdisp then
				call main_1_boardlist_2()
			else
				call main_1_boardlist_1()
			end if
		end if
		master_2=""
		rs.movenext
		if k<BoardCount then
			if rs("parentid")=0 and Board_count<ColSpanNum and simdisp then
				for i=Board_Count+1 to ColSpanNum
					response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
				next
				response.write "</tr>"
			end if
			if rs("parentid")=0 then
				response.write "</table><BR>"
				simdisp=false
				ColSpanNum=2
			end if
		else
			if Board_count<ColSpanNum and simdisp then
				for i=Board_Count+1 to ColSpanNum
					response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
				next
				response.write "</tr>"
			end if
		end if
	loop
	set rs=nothing
	response.write "</table>"
	response.write "<br>"
end sub

sub main_1_boardlist_1()
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
		response.write "<TR><TD align=middle width=100% class=tablebody1><table width=""100%"" cellspacing=0 cellpadding=0>"
		LastPostInfo=split(rs("LastPost"),"$")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		'=========================BoardInfo============================
		response.write "<TR><TD align=middle width=46 class=tablebody1>"
		dim LastViewBoard
		if request.cookies("LastView")("Board_"&rs(0))="" then
			if isnull(LastViewTime) then
				LastViewBoard=CDate("2003-1-1")
			else
				LastViewBoard=LastViewTime
			end if
		else
			if isnull(LastViewTime) then
				LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs(0)))
			elseif LastViewTime<CDate(request.cookies("LastView")("Board_"&rs(0))) then
				LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs(0)))
			else
				LastViewBoard=LastViewTime
			end if
		end if
		if Cint(reBoard_Setting(0))=1 then
			if LastViewBoard<CDate(LastPostInfo(2)) then
				response.write "<img src="&Forum_info(7)&Forum_pic(15)&" width=32 height=32 alt=��������>"
			else
				response.write "<img src="&Forum_info(7)&Forum_pic(14)&" width=32 height=32 alt=��������>"
			end if
		else
			if LastViewBoard<CDate(LastPostInfo(2)) then
				response.write "<img src="&Forum_info(7)&Forum_pic(7)&" width=32 height=32 alt=��������>"
			else
				response.write "<img src="&Forum_info(7)&Forum_pic(6)&" width=32 height=32 alt=��������>"
			end if
		end if
		response.write "</TD>"
		response.write "<TD width=1 bgcolor="&Forum_body(27)&">"
		response.write "<TD vAlign=top width=* class=tablebody1>"
		response.write "<TABLE cellSpacing=0 cellPadding=2 width=100% border=0>"
		response.write "<tr>"
		response.write "<td class=tablebody1 width=*>"
		response.write "<table width=100% border=0 cellSpacing=0 cellPadding=2>"
		response.write "<tr>"
		response.write "<td width=150>"
		'============= ͬѧ¼ ��ʼ ==================
		if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
			if cint(right(rs("txlpd"),1))=1 then
				response.write "<a href='z_school_class.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
			else
				response.write "<a href='list.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
			end if
		else
			response.write "<a href='list.asp?boardid="&rs(0)&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
		end if
		'============= ͬѧ¼ ���� ==================
		if rs("child")>0 then
			response.write "(<font title=""��"&rs("child")&"��������̳"">"&rs("child")&")</font>"
		end if
		response.write "</td>"
		response.write "<td width=* align=left>"
		response.write "<a href=announce.asp?boardid="&rs("boardid")&"><img src=pic/fabiao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"����������""></a> "
		response.write "<a href=vote.asp?boardid="&rs("boardid")&"><img src=pic/toupiao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"������ͶƱ""></a> "
		response.write "<a href=SmallPaper.asp?boardid="&rs("boardid")&"><img src=pic/xiaozibao.gif border=0 width=12 height=13 alt=""��"&rs("boardtype")&"����С�ֱ�""></a> "
		response.write "<a href=queryresult.asp?boardid="&rs("boardid")&"&stype=3><img src=pic/findposter.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"������""></a> "
		response.write "<a href=hotlist.asp?boardid="&rs("boardid")&"&stype=2&searchdate=30><img src=pic/retie.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"�����Ż���""></a> "
		response.write "<a href=elist.asp?boardid="&rs("boardid")&"><img src=pic/jinhua.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"�ľ�������""></a> "
		response.write "<a href=bbseven.asp?boardid="&rs("boardid")&"><img src=pic/chazhao.gif border=0 width=12 height=13 alt=""�鿴"&rs("boardtype")&"���¼�""></a>&nbsp;<font color="&forum_body(27)&" title='�����Ա:"&online(rs("boardid"),true)&"���������:"&guest(rs("boardid"),true)&"'>(�������� <b>"&online(rs("boardid"),true)+guest(rs("boardid"),true)&"</b> ��)</font>"
		response.write "</td>"
		response.write "</tr></table>"
		response.write "</td>"
		response.write "<td width=40 rowspan=2 align=center class=tablebody1>"
		if rs("indeximg")<>"" then
			'============= ͬѧ¼ ��ʼ ==================
			if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
				if cint(right(rs("txlpd"),1))=1 then
					response.write "<table align=left><tr><td align=right><a href=z_school_class.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
					if cint(forum_setting(76))=1 then
						response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
					end if
					response.write "></a></td><td width=20></td></tr></table>"
				else
					response.write "<table align=left><tr><td align=right><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
					if cint(forum_setting(76))=1 then
						response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
					end if
					response.write "></a></td><td width=20></td></tr></table>"
				end if
			else
				response.write "<table align=left><tr><td align=right><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0"
				if cint(forum_setting(76))=1 then
					response.write " style=""FILTER: alpha(opacity=40)"" onMouseOut=nereidFade(this,40,10,10) onMouseOver=nereidFade(this,100,0,10)"
				end if
				response.write "></a></td><td width=20></td></tr></table>"
			end if
			'============= ͬѧ¼ ��ʼ ==================
		end if
		response.write "</td><td width=200 rowspan=2 class=tablebody1>"
		if Cint(reBoard_Setting(2))=1 then
			response.write "��֤��̳������֤�û��������"
		else
			if isnull(LastPostInfo(7)) or LastPostInfo(7)="" then
				response.write "���⣺û���κ�����<BR>���ߣ���<BR>���ڣ�" & FormatDateTime(now(),1) & "��" & FormatDateTime(now(),4)&"&nbsp;<IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" alt=""û���κ�����"">"
			else
				response.write "���⣺<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"..</a><BR>���ߣ�<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a><BR>���ڣ�" & FormatDateTime(LastPostInfo(2),1) & "��" & FormatDateTime(LastPostInfo(2),4)&"&nbsp;<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" title=""ת����"&htmlencode(LastPostInfo(3))&"""></a>"
			end if
		end if
		response.write "</TD></TR><TR><TD width=*><FONT face=Arial><img src="&Forum_info(7)&Forum_pic(8)&" align=middle> "
		response.write rs("readme")
		response.write"</FONT></TD>"
		response.write "</TR><TR><TD class=tablebody2 height=20 width=*>������"&master_2&""
		response.write "</TD><td width=20 class=tablebody2></td><TD vAlign=middle class=tablebody2 width=200><table width=100% border=0><tr>"
		response.write "<td width=25% vAlign=middle><img src="&Forum_info(7)&Forum_pic(9)&" alt=������ align=absmiddle>&nbsp;<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=30% vAlign=middle><img src="&Forum_info(7)&Forum_pic(11)&" alt=���� border=0  align=absmiddle>&nbsp;"&rs("LastTopicNum")&"</td><td width=45% vAlign=middle><img src="&Forum_info(7)&Forum_pic(10)&" alt=���� border=0 align=absmiddle>&nbsp;"&rs("LastBBSNum")&"</td></tr>"
		response.write "</table></TD></TR></TABLE></td></tr></table></td></tr>"
		'============================End===============================
	end if
end sub

sub main_1_boardlist_2()
	TdTableWidth=100/ColSpanNum
	if Board_Count=ColSpanNum or Board_Count=0 then response.write "<tr>"
	response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
		response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>������"&master_2&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
		if rs("child")>0 then
			response.write "(<font title=""��"&rs("child")&"��������̳"">"&rs("child")&"</font>)"
		end if
		response.write "</td></tr><tr><td width=""50%"">���գ�<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=""50%"">������"&rs("LastBbsNum")&"</td></tr></table>"
	end if
	response.write "</td>"
	if Board_Count=ColSpanNum-1 then response.write "</tr>"
	if Cint(Board_Count)>Cint(ColSpanNum)-1 then 
		Board_Count=1
	else
		Board_Count=Board_Count+1
	end if
end sub

sub main_2()
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>"
	response.write "<TR>"
	response.write "<Th vAlign=center noWrap align=middle width=35 height=25>״̬</Th>"
	response.write "<Th vAlign=center noWrap align=middle width=*>��̳����</Th>"
	response.write "<Th vAlign=center noWrap align=middle width=80>����</Th>"
	response.write "<Th vAlign=center noWrap align=middle width=38>����</Th>"
	response.write "<Th vAlign=center noWrap align=middle width=38>����</Th>"
	response.write "<Th vAlign=center noWrap align=middle width=168>��󷢱�</Th>"
	response.write "</TR>"
	'��ʾ�ҵ��ղ���̳�б�ʼ
	call showmylist()
	if request.Cookies("bbslist")("Isshow")<>"0" then 
		'��ʾ�ҵ��ղ���̳�б����
		Dim cBoardID
		Response.Cookies("BoardList").expires= date+7
		cBoardID= Request("cBoardid")
		if isnumeric(cBoardID) then
			if request("CatLog")="N" then
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
			else
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
			end if
		end if
		dim simdisp
		simdisp=false
		k=0
		'ֻ�г����С��2�İ���
		sql="select * from board where depth<="&forum_setting(5)&" order by rootid,orders"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if not (rs.eof and rs.bof) then 
			BoardCount=rs.recordcount 
		else 
			boardcount=0 
		end if
		do while not rs.eof
			k=k+1
			reBoard_Setting=split(rs("Board_setting"),",")
			if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
				master_1=split(rs("BoardMaster"), "|")
				for i = 0 to ubound(master_1)
					if i>6 then
						master_2=master_2
					else
						master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
					end if
				next
				if i>2 then master_2=master_2 & "<font color=gray>More...</font>"
			else
				master_2="����"
			end if
			if rs("ParentID")=0 then
				if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
					ColSpanNum=reBoard_Setting(41)
					if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
					simdisp=true
				elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
					simdisp=false
					ColSpanNum=6
				else
					if Cint(reBoard_Setting(39))=1 then
						ColSpanNum=reBoard_Setting(41)
						if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
						simdisp=true
					else
						simdisp=false
						ColSpanNum=6
					end if
				end if
				response.write "<TR><Th colSpan=6 height=25 align=left id=TableTitleLink>&nbsp;"
				if simdisp then
					response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=Y"" title=""չ����̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>"
				else
					response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=N"" title=""�ر���̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=���뱾������̳>"&rs("boardtype")&"</a>"
				end if
				response.write "</Th></TR>"
				Board_Count=0
			else
				if simdisp then
					call main_2_boardlist_2()
				else
					call main_2_boardlist_1()
				end if
			end if
			master_2=""
			rs.movenext
			if k<BoardCount then
				if rs("parentid")=0 and Board_count<ColSpanNum and simdisp then
					for i=Board_Count+1 to ColSpanNum
						response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
					next
					response.write "</tr></table></td></tr>"
				end if
				if rs("parentid")=0 then
					simdisp=false
					ColSpanNum=6
				end if
			else
				if Board_count<ColSpanNum and simdisp then
					for i=Board_Count+1 to ColSpanNum
						response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
					next
					response.write "</tr></table></td></tr>"
				end if
			end if
		loop
		set rs=nothing
		response.write "</table>"
		'��ʾ�ҵ��ղ���̳�б�ʼ
	end if
	'��ʾ�ҵ��ղ���̳�б����
	response.write "<br>"
end sub

sub main_2_boardlist_1()
	dim boardstat
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
		LastPostInfo=split(rs("LastPost"),"$")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		'=========================BoardInfo============================
		response.write "<TR><TD width=26 class=tablebody2 valign=middle align=center>"
		dim LastViewBoard
		if request.cookies("LastView")("Board_"&rs("boardid"))="" then
			if isnull(LastViewTime) then
				LastViewBoard=CDate("2003-1-1")
			else
				LastViewBoard=LastViewTime
			end if
		else
			if isnull(LastViewTime) then
				LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs("boardid")))
			elseif LastViewTime<CDate(request.cookies("LastView")("Board_"&rs("boardid"))) then
				LastViewBoard=CDate(request.cookies("LastView")("Board_"&rs("boardid")))
			else
				LastViewBoard=LastViewTime
			end if
		end if
		if Cint(reBoard_Setting(0))=1 then
			if LastViewBoard<CDate(LastPostInfo(2)) then
				response.write "<img src="&Forum_info(7)&Forum_pic(15)&" width=32 height=32 alt=��������>"
			else
				response.write "<img src="&Forum_info(7)&Forum_pic(14)&" width=32 height=32 alt=��������>"
			end if
		else
			if LastViewBoard<CDate(LastPostInfo(2)) then
				response.write "<img src="&Forum_info(7)&Forum_pic(7)&" width=32 height=32 alt=��������>"
			else
				response.write "<img src="&Forum_info(7)&Forum_pic(6)&" width=32 height=32 alt=��������>"
			end if
		end if
		response.write "</TD><TD vAlign=top width=* class=tablebody1>"
		response.write "<table width=100% cellpadding=2 cellspacing=0>"
		response.write "<tr><td class=tablebody1 width=30% >"
		'============= ͬѧ¼ ��ʼ ==================
		if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
			if cint(right(rs("txlpd"),1))=1 then
				response.write"<a href='z_school_class.asp?boardid="&rs("boardid")&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
			else
				response.write "<a href='list.asp?boardid="&rs("boardid")&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
			end if
		else
			response.write "<a href='list.asp?boardid="&rs("boardid")&"'><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
		end if
		'============= ͬѧ¼ ���� ==================
		response.write "</td>"
		response.write "<td align=left class=tablebody1>"
		response.write "<a href=announce.asp?boardid="&rs("boardid")&"><img src=pic/fabiao.gif border=0 alt=""��"&rs("boardtype")&"����������""></a> "
		response.write "<a href=vote.asp?boardid="&rs("boardid")&"><img src=pic/toupiao.gif border=0 alt=""��"&rs("boardtype")&"������ͶƱ""></a> "
		response.write "<a href=SmallPaper.asp?boardid="&rs("boardid")&"><img src=pic/xiaozibao.gif border=0 alt=""��"&rs("boardtype")&"����С�ֱ�""></a> "
		response.write "<a href=queryresult.asp?boardid="&rs("boardid")&"&stype=3><img src=pic/findposter.gif border=0 alt=""�鿴"&rs("boardtype")&"������""></a> "
		response.write "<a href=hotlist.asp?boardid="&rs("boardid")&"&stype=2&searchdate=30><img src=pic/retie.gif border=0 alt=""�鿴"&rs("boardtype")&"�����Ż���""></a> "
		response.write "<a href=elist.asp?boardid="&rs("boardid")&"><img src=pic/jinhua.gif border=0 alt=""�鿴"&rs("boardtype")&"�ľ�������""></a> "
		response.write "<a href=bbseven.asp?boardid="&rs("boardid")&"><img src=pic/chazhao.gif border=0 alt=""�鿴"&rs("boardtype")&"���¼�""></a>&nbsp;<font color="&forum_body(27)&" title='�����Ա:"&online(rs("boardid"),true)&"���������:"&guest(rs("boardid"),true)&"'>(�������� <b>"&online(rs("boardid"),true)+guest(rs("boardid"),false)&"</b> ��)</font>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr><td class=tablebody1 colspan=2 width=100% >"
		if rs("indeximg")<>"" then
			'============= ͬѧ¼ ��ʼ ==================
			if not isnull(rs("txlpd")) and rs("txlpd")<>"" then
				if cint(right(rs("txlpd"),1))=1 then
					response.write "<table align=left><tr><td><a href=z_school_class.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
				else
					response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
				end if
			else
				response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
			end if
			'============= ͬѧ¼ ���� ==================
		end if
		response.write rs("readme")
		response.write "</td></tr></table>"
		response.write "</TD><TD vAlign=center align=middle class=tablebody2 width=80>"
		response.write master_2
		response.write "</TD>"
		response.write "<TD vAlign=center noWrap align=middle width=38 class=tablebody1>"&rs("LastTopicNum")&"</TD>"
		response.write "<TD vAlign=center noWrap align=middle width=38 class=tablebody1>"&rs("LastBbsNum")&"</TD>"
		response.write "<TD noWrap width=168 class=tablebody2>"
		if isnull(LastPostInfo(7)) or LastPostInfo(7)="" then
			response.write "���ߣ��ޣ�<IMG border=0 src=""&Forum_info(7)&Forum_statePic(5)&"" alt=""û���κ�����""><br>"
			response.write "���⣺û���κ�����<br>"
			response.write FormatDateTime(now(),1)&"��"&FormatDateTime(now(),4)
		else
			response.write "���ߣ�<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a>��<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" title=""ת����"&htmlencode(LastPostInfo(3))&"""></a><br>"
			response.write "���⣺<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"...</a><br>"
			response.write FormatDateTime(LastPostInfo(2),1)&"��"&FormatDateTime(LastPostInfo(2),4)
		end if
		response.write "</TD></TR>"
	end if
end sub

sub main_2_boardlist_2()
	TdTableWidth=100/ColSpanNum
	if Board_Count=0 then response.write "<tr><td colspan=6 class=tablebody1 width=""100%""><table width=""100%""><tr>"
	if Board_Count=ColSpanNum then response.write "<tr>"
	response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
		response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>������"&master_2&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
		if rs("child")>0 then
			response.write "(<font title=""��"&rs("child")&"��������̳"">"&rs("child")&"</font>)"
		end if
		response.write "</td></tr><tr><td width=""50%"">���գ�<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=""50%"">������"&rs("LastBbsNum")&"</td></tr></table>"
	end if
	response.write "</td>"
	if Board_Count=ColSpanNum-1 then response.write "</tr>"
	if Cint(Board_Count)>Cint(ColSpanNum)-1 then 
		Board_Count=1
	else
		Board_Count=Board_Count+1
	end if
end sub

'��ʾ�ҵ��ղ���̳�б�ʼ
Sub showmylist()
	Dim cBoardID
	Response.Cookies("BoardList").expires= date+7
	cBoardID= Request("cBoardid")
	if isnumeric(cBoardID) then
		if request("CatLog")="N" then
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
		else
			Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
		end if
	end if
	Dim boardstat
	dim boardids
	Set rs= Server.CreateObject("ADODB.Recordset")
	if isnull(request.cookies("bbslist")("mylist")) or request.cookies("bbslist")("mylist")="" then
		boardids =0
	else
		boardids=trim(request.cookies("bbslist")("mylist"))
	end if
	if cint(request("setup"))=1 then
		If BoardMaster Or Master Then
			Rs.Open "select boardid,BoardType from board order by rootid ,orders", Conn, 0, 1
		Else
			Rs.Open "select boardid,BoardType from board order by rootid ,orders", Conn, 0, 1
		End If
		dim K
		k=1
		Response.Write "<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>"&vbNewLine
		Response.Write "<form action=""cookies.asp?action=mylist"" method=post>"&vbNewLine
		Response.Write "<tr>"&vbNewLine
		Response.Write"<th height=""25"" colspan=6 >��������ϲ������̳�б�</th>"&vbNewLine
		Response.Write "</tr>"&vbNewLine&"<tr align=left>"&vbNewLine
		do while not rs.eof 
			if k>6 then
				Response.Write "</tr>"&vbNewLine&"<tr align=left>"&vbNewLine
				k =1
			end if
			Response.Write"<td class=tablebody1>"
			Response.Write "<input type=""checkbox"" name=""mylist"" value="""&rs(0)&""" "
			if instr(","&boardids&",",","&rs(0)&",")>0 then 
				Response.Write " checked "
			end if
			Response.Write ">"&rs(1)&"</td>"&vbNewLine
			k= k+1  
			rs.movenext
		loop
		rs.close
		set rs=nothing
		if k<6 then
			do while k<=6
				Response.Write "<td class=tablebody1>&nbsp;</td>"&vbNewLine
				k=k+1
			loop
		end if
		Response.Write "</tr>"&vbNewLine&"<tr align=center>"&vbNewLine
		Response.Write "<TD class=tablebody1 colspan=3><input type=""radio"" name=""IsShow"" value="""" "
		if request.cookies("bbslist")("isshow")<> "0" then
			Response.Write "checked"
		end if
		Response.Write ">��ʾȫ����̳�б�</td>"&vbNewLine&"<TD class=tablebody1 colspan=3>"
		Response.Write "<input type=""radio"" name=""IsShow"" value=""0"" "
		if request.cookies("bbslist")("isshow")= "0" then
			Response.Write "checked"
		end if
		Response.Write ">����ʾȫ����̳�б�</td>"&vbNewLine
		Response.Write "</tr>"&vbNewLine&"<tr align=center>"&vbNewLine&"<TD class=tablebody1 colspan=3>"
		Response.Write "<input type=""submit"" name=""Submit"" value=""�� ��""></td>"&vbNewLine
		Response.Write "<TD class=tablebody1 colspan=3><a href="&scriptname&">[�� ��]</a></td>"&vbNewLine
		Response.Write "</tr>"&vbNewLine
		Response.Write("</form>")&vbNewLine
		Response.Write  "</table>"&vbNewLine
	else
		dim simdisp
		k=0
		simdisp=false
		if Cint(Forum_setting(4))=1 then
			if request.cookies("BoardList")("0BoardID")="NotShow" then
				ColSpanNum=2
				simdisp=true
			else
				simdisp=false
				ColSpanNum=6
			end if
			response.write "<TR><Th colSpan=6 height=25 align=left id=TableTitleLink>&nbsp;"
			if simdisp then
				response.write "<a href=""?cBoardid=0&Catlog=Y"" title=""չ����̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a>����ϲ������̳  <a href=""?setup=1"">[����]</a>"
			else
				response.write "<a href=""?cBoardid=0&Catlog=N"" title=""�ر���̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a>����ϲ������̳  <a href=""?setup=1"">[����]</a>"
			end if
			if request.cookies("bbslist")("isshow")= "0" then
				Response.Write "  <a href=""cookies.asp?action=show0"">[��ʾȫ����̳�б�]</a>"
			else
				Response.Write "  <a href=""cookies.asp?action=show1"">[����ȫ����̳�б�]</a>"
			end if
			response.write "</Th></TR>"
			'ֻ�г����С��2�İ���
			sql="select * from board where Boardid in ("&boardids&") order by rootid ,orders"
			Rs.Open sql, Conn, 1, 1
			if not (rs.eof and rs.bof) then BoardCount=rs.recordcount else boardcount=0 end if
			if rs.eof and rs.bof then
				Response.write "<tr><td colspan=6 class=tablebody1 align=center height=25>Ŀǰ��ϲ������̳�б��ǿյ�,�����Ե�� <a href=?setup=1>[����]</a> ��ӻ��޸����ϲ������̳"
				Response.write "</td></tr>"
			else
				Board_Count=0
				do while not rs.eof
					k=k+1
					reBoard_Setting=split(rs("Board_setting"),",")
					if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
						master_1=split(rs("BoardMaster"), "|")
						for i = 0 to ubound(master_1)
							if i>6 then
								master_2=master_2
							else
								master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
							end if
						next
						if i>2 then master_2=master_2 & "<font color=gray>More...</font>"
					else
						master_2="����"
					end if
					if simdisp then
						call main_2_boardlist_2()
					else
						call main_2_boardlist_1()
					end if
					master_2=""
					rs.movenext
					if k>=BoardCount then
						if Board_count<ColSpanNum and simdisp then
							for i=Board_Count+1 to ColSpanNum
								response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
							next
							response.write "</tr></table></td></tr>"
						end if
					end if
				loop
			end if
			rs.close
			if request.Cookies("bbslist")("Isshow")="0" then 
				response.write "</table>"
			end if
		else
			if request.cookies("BoardList")("0BoardID")="NotShow" then
				ColSpanNum=4
				simdisp=true
			else
				simdisp=false
				ColSpanNum=2
			end if
			response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
			response.write "<TR><Th colSpan="&ColSpanNum&" height=25 align=left id=TableTitleLink>&nbsp;"
			if simdisp then
				response.write "<a href=""?cBoardid=0&Catlog=Y"" title=""չ����̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a>����ϲ������̳  <a href=""?setup=1"">[����]</a>"
			else
				response.write "<a href=""?cBoardid=0&Catlog=N"" title=""�ر���̳�б�""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a>����ϲ������̳  <a href=""?setup=1"">[����]</a>"
			end if
			if request.cookies("bbslist")("isshow")= "0" then
				Response.Write "  <a href=""cookies.asp?action=show0"">[��ʾȫ����̳�б�]</a>"
			else
				Response.Write "  <a href=""cookies.asp?action=show1"">[����ȫ����̳�б�]</a>"
			end if
			response.write "</Th></TR>" & vbNewLine
			'ֻ�г����С��2�İ���
			sql="select * from board where Boardid in ("&boardids&") order by rootid ,orders"
			Rs.Open sql, Conn, 1, 1
			if not (rs.eof and rs.bof) then BoardCount=rs.recordcount else boardcount=0 end if
			if rs.eof and rs.bof then
				Response.write "<tr><td colspan=2 class=tablebody1 align=center height=25>Ŀǰ��ϲ������̳�б��ǿյ�,�����Ե�� <a href=?setup=1>[����]</a> ��ӻ��޸����ϲ������̳"
				Response.write "</td></tr>"
			else
				Board_Count=0
				do while not rs.eof
					k=k+1
					reBoard_Setting=split(rs("Board_setting"),",")
					if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
						master_1=split(rs("BoardMaster"), "|")
						for i = 0 to ubound(master_1)
							if i>6 then
								master_2=master_2
							else
								master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
							end if
						next
						if i>2 then master_2=master_2 & "<font color=gray>More...</font>"
					else
						master_2="����"
					end if
					if simdisp then
						call main_1_boardlist_2()
					else
						call main_1_boardlist_1()
					end if
					master_2=""
					rs.movenext
					if k>=BoardCount then
						if Board_count<ColSpanNum and simdisp then
							for i=Board_Count+1 to ColSpanNum
								response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
							next
							response.write "</tr>"
						end if
					end if
				loop
			end if
			rs.close
			response.write "</table>"
		end if
	end if
	set rs=nothing
	response.write "<BR>"
end sub
'��ʾ�ҵ��ղ���̳�б����
%>
<!--#include file="z_mingxing.asp"-->
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->