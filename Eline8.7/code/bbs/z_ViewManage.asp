<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_ViewManage_Conn.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
dim page
if request("action")="reset" and master and membername<>"" then
	connvm.execute("delete * from [fuser]")
	Response.Redirect "z_ViewManage.asp"
end if
If Request("page") = "" or not isInteger(Request("page")) then
	page = 1
Else
	page = CINT(Request("page"))
End If
if not founduser then
	stats="��������"
	call nav()
	call head_var(2,0,"","")
	Errmsg=Errmsg+"<br>"+"<li>��û�в鿴����������Ȩ�ޣ����ȵ�¼����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	if request("action")="hy" then
		stats="��Ա���"
		call nav()
		call head_var(2,0,"","")
		call hy()
	else
		stats="��������"
		call nav()
		call head_var(2,0,"","")
		call bz()	
	end if 
end if
call activeonline()
call footer() 

sub bz()
	dim Gzdays,Gzlbt,Gzhmon,Gzwealth,Gzuserep,Gzusercp,Gzpower,Gzifmess,Gzjbgz,pagenum,ifAnn,ifAutoCancelMSG
	dim gzifarticle,gzsuperfj,gzparticle,gzplogins,gzplogs,gziflogs,gziflogins,ifautocancel,autocday,gzmasterfj
	Gzdays=7		'�����ռ����ÿ7�췢��һ��
	gzifmess="1"		'��н�Ƿ����֪ͨ����  1�ǣ�0��
	pagenum=8		'ÿҳ�������г�������
	
	Gzjbgz=100		'������ÿ���������
	gzsuperfj=500		'�����������ӹ���
	gzmasterfj=1000		'�����������ӹ���
	ifautocancel="1"	'�Ƿ��Զ�ȡ��ʧְ����ְ��  1�Զ���0�ֶ�
	ifAnn="1"		'�Ƿ񷢲�����  1�ǣ�0��
	ifAutoCancelMSG="1"		'�Ƿ����ѱ�ȡ���ʸ�İ�����������֪ͨ  1�ǣ�0��
	autocday=15		'����������δ��¼ȡ������ְ��
	
	gzifarticle="1"		'���ּ��㹫ʽ�Ƿ���㷢������1���㣬0������
	gzparticle=0.5		'����������ϵ������100���������빤������50
	gziflogins="1"		'���ּ��㹫ʽ�Ƿ�����¼������1���㣬0������
	gzplogins=0.5		'��¼��������ϵ�������¼100*0.5��ʵ�ʼ���50
	gziflogs="1"		'���ּ��㹫ʽ�Ƿ����ά������1���㣬0������
	gzplogs=1		'ά��������ϵ����Ĭ��Ϊ1
	
	Gzlbt=20		'�õ�ҵ���������蹤�����ֵ���
	gzhmon=20		'����ÿ���Ӷ�����һ��ҵ������
	gzwealth=100		'һ��ҵ�����ʼӶ���Ǯ
	gzuserep=50		'һ��ҵ�����ʼӶ��پ���
	gzusercp=20		'һ��ҵ�����ʼӶ�������
	gzpower=1		'һ��ҵ�����ʼӶ�������
  
	'------------�������ʵ�������޸����ϲ���-----
	dim board,rsu,UserName,Absentday,article,Message,RsB,Bcount,BoardName,allbm,RsL,RsF,Rstmp,p
	dim Bgzday,Btgz,savemoney,saveep,savecp,savepower
	dim Bnumgz,Bnumlog,saveMessage,masterlist,Barticle
	response.write "<table cellpadding=3 cellspacing=1 class=tableborder1 align=center>"
	response.write "<tr>"
	response.write "<th valign=middle colspan=7 align=center height=25>�����������</th>"
	response.write "</tr>"
	response.write "<tr>"
	response.write "<td colspan=7 class=tablebody1><br>"
	response.write "������л��λ������ <font color="&forum_body(8)&">"&Forum_info(0)&"</font> �Ĵ���֧�֣�"
	response.write "��վÿ <fotn color="&forum_body(8)&">"&gzdays&"</font> ��Ը�λ�������л��ֽ�����<br><br>"
	response.write "�����������ֵļ��㹫ʽΪ��"
	response.write "���� <b>"&gzdays&"</b> ����"
	if gziflogins="1" then 
    response.write "��¼���� <b>*</b> <b>"&formatnumber(gzplogins,1,-1)&"</b>"
    if gzifarticle="1" or gziflogs="1" then response.write " <b>+</b> "
	end if
	if gzifarticle="1" then 
		response.write "���ڷ����� <b>*</b> <b>"&formatnumber(gzparticle,1,-1)&"</b>"
		if gziflogs="1" then response.write " <b>+</b> "
	end if
	if gziflogs="1" then response.write "����ά������ <b>*</b> <b>"&formatnumber(gzplogs,1,-1)&"</b>"
	response.write "��<br>"
	response.write "����ʵ�ʸ�������Ϊ�������ʼ���ҵ�����ʣ�ά������Ϊ��İ�������ͣн��<br>"
	response.write "���������Ļ�������Ϊÿ�� <font color="&forum_body(8)&"><b>"&Gzjbgz&"</b></font> Ԫ�������������� <font color="&forum_body(8)&"><b>"&gzsuperfj&"</b></font> Ԫ���ӽ���������Ա���� <font color="&forum_body(8)&"><b>"&gzmasterfj&"</b></font> Ԫ���ӽ�����<br>"
	response.write "�����������ֳ��� <font color="&forum_body(8)&"><b>"&Gzlbt&"</b></font> ������ҵ�����ʣ�ҵ�����ʼ��㷽����<br>"
	response.write "���������Ĺ�������ÿ��� <font color="&forum_body(8)&"><b>"&gzhmon&"</b></font> �㣬"
	response.write "�����ֽ� <font color="&forum_body(8)&"><b>"&gzwealth&"</b></font> Ԫ��"
	response.write "���� <font color="&forum_body(8)&"><b>"&gzpower&"</b></font> �㣬"
	response.write "���� <font color="&forum_body(8)&"><b>"&gzuserep&"</b></font> �㣬"
	response.write "���� <font color="&forum_body(8)&"><b>"&gzusercp&"</b></font> �㡣<br>"
	response.write "�������λ����ע�⣬���� <font color="&forum_body(8)&"><b>"&autocday&"</b></font> ��δ��¼�İ��������Զ�ȡ�������ʸ�<br><br>"
	response.write "������л��ҵ�֧�֣�<br><br>"
	response.write "</td>"
	response.write "</tr>"
	response.write "<tr height=25>"
	response.write "<th><B>����</b></th>"	
  response.write "<th><B>���ΰ��</b></th>"
	response.write "<th><B>����������</b></th>"
	response.write "<th><B>��н/��¼��</b></th>"
	response.write "<th><B>��¼</b></th>"
	response.write "<th><B>ά����/����</b></th>"
	response.write "<th><B>Ԥ�ڹ���/����</b></th>"
	response.write "</tr>"
	sql="select UserName,UserGroupID,Article,lastlogin,logins from [user] where UserGroupID < 4 order by lastlogin desc"
	set RsU=Server.createobject("ADODB.Recordset")
	RsU.open sql,conn,1,1
	if rsu.eof or rsu.bof then
		response.write "<tr><td class=tablebody1 align=center colspan=7 height=30>��̳��û���κΰ�����</td></tr>"
	else
		Rsu.PageSize = Pagenum
		if page<1 then page=1
		If page>RsU.Pagecount Then page=RsU.Pagecount              
		rsu.AbsolutePage=page
		set RSB=Server.createobject("ADODB.Recordset")
		allbm=""
		do while not RsU.eof
			UserName=RsU("UserName")
			if isnull(RsU("lastlogin")) then
				AbsentDay=autocday+1
			else
				AbsentDay=DateDiff("d",RsU("lastlogin"),Now)
			end if
			article=RsU("article")
			if AbsentDay =< autocday/5  then
				Message="<font face=Wingdings color=blue>v</font> <font color=#0080d5>�����Ƚ����渺��</font>"
			elseif AbsentDay =< autocday*2/5  then
				Message="����ɢ���� <font color="&forum_body(8)&"><b>"&int(autocday*2/5)&"</b></font> �����е�¼��"
			elseif AbsentDay =< autocday*3/5 then
				Message="ע�⣬�Ѿ� <font color="&forum_body(8)&"><b>"&int(autocday*2/5)&"</b></font> ������û�е�¼�ˣ�</font>"
			elseif AbsentDay =< autocday*4/5 then
				Message="����ʧְ���Ѿ� <font color="&forum_body(8)&"><b>"&int(autocday*3/5)&"</b></font> ������û�е�¼�ˣ�</font>"
			elseif AbsentDay =< autocday-1 then
				Message="����ʧְ���Ѿ� <font color="&forum_body(8)&"><b>"&int(autocday*4/5)&"</b></font> ������û�е�¼�ˣ������ǳ���ְ��</font>"
			else
				Message="<font color="&forum_body(8)&"><b>�����������ְ���Ѿ�������ְ��</b></font>"
				if allbm="" then
					allbm=UserName
				else
					allbm=allbm&"��"&UserName
				end if
		    if ifautocancel="1" then
					connvm.execute("delete from fuser where username='"&UserName&"'")
					dim rs5,BoardMaster_1,BoardMaster_2,k
					dim oldMinArticle
					oldMinArticle=99999999
					set rs5=conn.execute("select * from usertitle order by MinArticle desc")
					do while not rs5.eof
						conn.execute("update [user] set usergroupid=4,userclass='"&rs5("usertitle")&"',titlepic='"&rs5("titlepic")&"' where username='"&UserName&"' and (article<"&oldMinArticle&" and article>="&rs5("MinArticle")&" )")
						oldMinArticle=rs5("MinArticle")
						rs5.movenext
					loop
					rs5.close
					set rs5=conn.execute("select BoardMaster,boardid from board where instr('|'&BoardMaster&'|','|"&UserName&"|')>0")
					do while not rs5.eof
						BoardMaster_2=""
						k=0
						BoardMaster_1=split(rs5(0),"|") 
						for i=0 to ubound(BoardMaster_1)
							if lcase(trim(UserName))=lcase(trim(BoardMaster_1(i))) then 
								BoardMaster_1(i)="" 
							end if
							if trim(BoardMaster_1(i))<>"" then
								if BoardMaster_2="" then
									BoardMaster_2=BoardMaster_1(i)
								else
									BoardMaster_2=BoardMaster_2&"|"&trim(BoardMaster_1(i))
								end if
							end if
						next
						conn.execute("update [board] set BoardMaster='"&BoardMaster_2&"' where boardid="&rs5(1)&"")
						rs5.movenext
					loop
					set rs5=nothing
					if ifAutoCancelMSG then
						conn.execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&Username&"','"&forum_info(0)&"','ϵͳ��Ϣ�����İ����ʸ���ȡ��','ԭ�򣺳���"&autocday&"��û�е�¼ϵͳ����������ʧְ����ڶ��ӣ�',Now(),0,1)")
					end if
				end if
			end if
			BCount=0
			set RsB=conn.execute("select BoardType,BoardMaster,BoardID from [board] where instr('|'&BoardMaster&'|','|"&UserName&"|')>0")
			if not RsB.eof then
				BoardName=""
				do while not RsB.eof
					if BoardName="" then
						BoardName="<a href=list.asp?boardid="&RsB("BoardID")&" target=_blank>"&RsB("BoardType")&"</a>"
					else
						BoardName=BoardName&"<br>"&"<a href=list.asp?boardid="&RsB("BoardID")&" target=_blank>"&RsB("BoardType")&"</a>"
					end if
					Bcount=Bcount+1
					RsB.movenext
				loop
			else 
				BoardName="<font color="&forum_body(8)&"><b>û�е��εİ��</b></font>"
			end if
			set RsF=connvm.execute("select * from [fuser] where username='"&Username&"' ")
			if RsF.eof then
				set Rstmp=server.createobject("adodb.recordset")
				Rstmp.open "select count(l_announceid) from [log] where l_username='"&UserName&"' group by l_announceid",conn,1,1
				connvm.execute("insert into fuser(username,FuLiDate,oldlogs,oldlogins,oldarticle) values('"&UserName&"',date()-1,"&Rstmp.recordcount&","&RsU("logins")&","&article&")")
				set Rstmp=nothing
			end if
			set Rsl=server.createobject("adodb.recordset")
			Rsl.open "select count(l_announceid) from [log] where l_username='"&UserName&"' group by l_announceid",conn,1,1
			if RsF.eof then
				Bnumgz=0
				Bnumlog=0
				Barticle=0
				Bgzday=(date()-1+gzdays)
			else
				Bnumgz=cint(Rsl.recordcount-Rsf("oldlogs"))
				Bnumlog=cint(RsU("logins")-Rsf("oldlogins"))
				Bgzday=(RsF("Fulidate")+gzdays)
				Barticle=(RsU("article")-Rsf("oldarticle"))
			end if
			'���ּ��㹫ʽ
			if gziflogins<>"1" then gzplogins=0
			if gzifarticle<>"1" then gzparticle=0
			if gziflogs<>"1" then gzplogs=0
			Btgz=(Bnumlog*gzplogins+Barticle*gzparticle+Bnumgz*gzplogs)
			if AbsentDay<=gzdays and Bnumgz>0 then
				if Btgz>Gzlbt then
					savemoney=(((Btgz-Gzlbt)\Gzhmon)*Gzwealth+gzjbgz*Bcount)
					savepower=(((Btgz-Gzlbt)\Gzhmon)*Gzpower)
					saveep=(((Btgz-Gzlbt)\Gzhmon)*Gzuserep)
					savecp=(((Btgz-Gzlbt)\Gzhmon)*Gzusercp)
				elseif Btgz>=0 then
					savemoney=gzjbgz*Bcount
					savepower=0
					saveep=0
					savecp=0
				end if
				if savemoney>0 and RsU("usergroupid")=2 then savemoney=cint(savemoney+gzsuperfj)
				if savemoney>0 and RsU("usergroupid")=1 then savemoney=cint(savemoney+gzmasterfj)
			else
			  savemoney=0
			  savepower=0
			  saveep=0
			  savecp=0
			end if
			if date()>=Bgzday and not RsF.eof and savemoney>0 then
			  savemessage="("&FormatDateTime(Rsf("Fulidate"),2)&"��"&FormatDateTime(Bgzday,2)&")����������("&cint(Btgz)&")����Ǯ("&savemoney&")������("&savepower&")������("&saveep&")������("&savecp&")"
			  connvm.execute("update [fuser] set Fulidate=date(),oldlogs="&Rsl.recordcount&",oldlogins="&RsU("logins")&",lastinfo='"&savemessage&"',oldarticle="&Rsu("article")&" where username='"&Username&"'")
			  conn.execute("update [user] set userwealth=userwealth+"&savemoney&",userep=userep+"&saveep&",usercp=usercp+"&savecp&",userpower=userpower+"&savepower&" where username='"&Username&"'")
			  if gzifmess="1" then
			    savemessage=savemessage&"�����ٴζ���Ϊ��վ����������ʾ��л������Ŭ����:)"
					conn.Execute("insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&Username&"','"&Forum_info(0)&"','ϵͳ��Ϣ�����Ĺ���','"&checkSTR(savemessage)&"',Now(),0,1)")
				end if
			end if
			response.write "<tr>"
			response.write "<td rowspan=2 class=tablebody1 align=left>&nbsp;"
			response.write "<a href=javascript:openScript('messanger.asp?action=new&touser="&UserName&"',500,400)>"
			select case Rsu("usergroupid")
			case 1
				response.write "<img src="&Forum_info(7)&Forum_pic(0)&" alt=��"&UserName&"����Ϣ height=18 border=0>"
			case 2
				response.write "<img src="&Forum_info(7)&Forum_pic(16)&" alt=��"&UserName&"����Ϣ height=18 border=0>"
			case 3
				response.write "<img src="&Forum_info(7)&Forum_pic(1)&" alt=��"&UserName&"����Ϣ height=18 border=0>"
			end select
			response.write "</a>&nbsp;"
			response.write "<a href=dispuser.asp?name="&UserName&" target=_blank title=""�鿴"&htmlencode(UserName)&"����Ϣ"">"
			if membername=username then
				response.write "<font color=blue>"
			end if
			response.write UserName
			if membername=username then
				response.write "</font>"
			end if
			response.write "</a>"
			response.write "</td>"
			response.write "<td rowspan=2 class=tablebody1 align=center>"
			response.write BoardName
			response.write "</td>"
			response.write "<td rowspan=2 class=tablebody1 align=center>���� <font color="&forum_body(8)&"><b>"&article&"</b></font> ƪ </td>"
			response.write "<td class=tablebody1>��н���ڣ�<b>"&FormatDateTime(Bgzday,2)&"</b></td>"
			response.write "<td class=tablebody1>���ڣ�<font color=blue><b>"&Bnumlog&"</b></font></td>"
			response.write "<td class=tablebody1>ά������<font color=blue><b>"&Bnumgz&"</font></td>"
			response.write "<td class=tablebody1 height=20 valign=top><img src=pic/mybbs.gif width=16 height=16 align=absmiddle alt=�ð������ڹ��ʣ�"
			if Rsf.eof then 
        response.write "0"
      else
        response.write Rsf("lastinfo")
      end if
			response.write ">&nbsp;<font color=blue>���֣�"
			response.write "<b><font color=red>"&cint(Btgz)&"</font></b> <font color=black>|</font> "
			if AbsentDay<=gzdays and Bnumgz>0 then  
				response.write "���ʣ���Ǯ(<b><font color="&forum_body(8)&">"&savemoney&"</font></b>)"
			else
				response.write " <b><font color="&forum_body(8)&">�˰������ڱ�ͣн</font></b>"
			end if
			if savepower>0 then response.write "����(<b><font color="&forum_body(8)&">"&savepower&"</font></b>)"
			if saveep>0 then response.write "����(<b><font color="&forum_body(8)&">"&saveep&"</font></b>)"
			if savecp>0 then response.write "����(<b><font color="&forum_body(8)&">"&savecp&"</font></b>)"
			response.write "</font></td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td class=tablebody1>����¼��<b>"&FormatDateTime(rsU("lastlogin"),2)&"</b></td>"
			response.write "<td  class=tablebody1>ʧ�٣�<font color=blue><b>"&AbsentDay&"</b></font></td>"
			response.write "<td class=tablebody1>��������<font color=blue><b>"&Barticle&"</b></font></td>"
			response.write "<td class=tablebody1><font color=#0080d5>������"&Message&"</font></td>"
			response.write "</tr>"
			p=p+1  
			RsF.close
			Set RsF=nothing
			RsL.close
			set Rsl=nothing
			RsU.movenext
  		if p>= rsu.PageSize then exit do
		loop
		response.write "<tr>"
		response.write "<td align=center height=25 class=tablebody2><a href=z_viewmanage.asp?action=hy><font color=blue>[��Ա���]</font></a></td>"
		response.write "<td height=25 colspan=6 class=tablebody2>"
		response.write "<marquee width=90% scrolldelay=100 scrollamount=4 onmouseout=""if (document.all!=null){this.start()}"" onmouseover=""if (document.all!=null){this.stop()}"">"
		if allbm<>"" then
			response.write allbm&"�Ѿ�����ʧְ��������������ְ��"
		else
			response.write "���а�������������<b>"&forum_info(0)&"</b>��л��ҵ�Ŭ����"
		end if
		response.write "</MARQUEE>"
		response.write "</td></tr></table>"
		response.write "<table cellpadding=0 cellspacing=0 width="&forum_body(12)&" align=center>"
		response.write "<tr>"
		response.write "<tr><td height=25 valign=middle colspan=5 align=left>&nbsp;&nbsp;��&nbsp;<b>"&Forum_info(0)&"</b> Ŀǰ���а��� <b>"&rsu.RecordCount&"</b> �ˣ��� <b>"&page&"</b> / <b>"&rsu.PageCount&"</b> ҳ</td>"
		response.write "<td height=25 valign=middle colspan=2 align=right>"
		if master then response.write "<a href=z_viewmanage.asp?action=reset>[<font color=red><b title=""ע�⣺ֻ����������˹�����־log��󣬲���Ҫ���д������"">�ؽ���������</b></font>]</a>&nbsp;"
		call disppagenum(page,rsu.pagecount,"?page=","")
		response.write "</td>"
		response.write "</tr>"
	end if
	response.write "</table>"
	rsU.close
	Set RsU=nothing
	if ifAnn and not allbm="" then
		conn.execute("insert into bbsnews(boardid,title,content,username,addtime) values (0,'�۹���ݳ���"&allbm&" ����һְ��','ԭ�򣺳��� "&autocday&" ��ûӵ�¼ϵͳ棬��������ʧְ����ڶ��ӣ�','"&forum_info(0)&"',now())")
	end if
end sub

sub hy()
	dim GlTime,sex
	dim PageCount
	PageCount=10
	set rs = server.CreateObject ("adodb.recordset")
	sql="select username,adddate,Sex,Userclass,UserGroup,lastlogin from [user] where logins <= 1 and Article=0 and datediff('d',lastlogin,date())>15 order by lastlogin"
	rs.open sql,conn,1,1
	response.write "<table Class=TableBorder1 cellspacing=1 cellpadding=3 align=center>"
	response.write "<tr>"
	response.write "<th height=25 align=center colspan=6>��Ա��¼���</td>"
	response.write "</tr>"
	if rs.eof or rs.bof then
		response.write "<tr><td colspan=6 class=tablebody1 height=30 align=center>���л�Աȫ���Ƚ����ģ�</td></tr>"
		response.write "<tr>"
		response.write "<td class=tablebody2 align=center height=25 colspan=6><a href=z_viewmanage.asp><font color=blue>[��������]</font></a></td>"
		response.write "</tr>"
	else
		response.write "<tr>"
		response.write "<td colspan=6 class=tablebody1><br>"
		response.write "������ӭ���� <font color="&forum_body(8)&">"&Forum_info(0)&"</font> ��<br><br>"
		response.write "���������г�����ֻ��¼��һ�Σ���û������κ����µĻ�Ա��������ô������Ϻ��������������ģ�<br>"
		response.write "����������������ֻ��һ�εĻ����벻Ҫע���ˣ�<br>"
		response.write "������������ʡ�����ݿ�Ŀռ䣬��ʡ�����ʱ�䣬������<br><br>"
		response.write "</td>"
		response.write "</tr>"
		response.write "<tr height=25>"
		response.write "<th>��Ա����</th>"
		response.write "<th>ע������</th>"
		response.write "<th>��ʧ����</th>"
		response.write "<th>��Ա�Ա�</th>"
		response.write "<th>��Ա����</th>"
		response.write "<th>������Ϣ</th>"
		response.write "</tr>"
		rs.AbsolutePage=Page
		rs.PageSize=PageCount
		do while not rs.EOF
			response.write "<tr height=25 align=center>"
			response.write "<td class=tablebody1 align=center><a href=dispuser.asp?name="&rs(0)&" target=""_blank"" title=""�鿴"&rs(0)&"�ĸ�������"">"&rs(0)&"</a></td>"
			response.write "<td class=tablebody1 align=center>"&rs(1)&"</td>"
			GlTime= datediff("d",rs(5),date())
			response.write "<td class=tablebody1 align=center><b><font color=red>"&GlTime&"</font></b> ��</td>"
			if rs(2)=1 then
				sex="��"
			elseif rs(2)=0 then
				sex="Ů" 
			end if
			response.write "<td class=tablebody1 align=center>"&sex&"</td>"
			response.write "<td class=tablebody1 align=center>"&rs(4)&"</td>"
			response.write "<td class=tablebody1 align=center><a target=""_blank"" href=messanger.asp?action=new&touser="&rs("username")&"><img src=pic/message.gif alt=�������� border=0></a></td>"
			response.write "</tr>" 
			i=i+1
			rs.movenext
			if i>= rs.PageSize then exit do
		loop  
		response.write "<tr>"
		response.write "<td class=tablebody2 align=center height=25><a href=z_viewmanage.asp><font color=blue>[��������]</font></a></td>"
		response.write "<td class=tablebody2>&nbsp;&nbsp;�� <b>"&rs.RecordCount&"</b> �� �� <b>"&page&"</b> / <b>"&rs.PageCount&"</b> ҳ</td>"
		response.write "<td class=tablebody2 align=right colspan=5>"
		call disppagenum(page,rs.pagecount,"?page=","&action=hy")
		response.write "</td>"
		response.write "</tr>"
	end if
	response.write "</table>"
	rs.Close
	set rs = nothing 
end sub
%>
