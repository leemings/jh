<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
	dim isdisp,bbseveninfo
	dim endpage
	dim totalrec
	dim n
	dim currentpage,page_count,Pcount
	dim bgcolor
	if boardid=0 then
	stats="��̳���¼��б�"
	call nav()
	call head_var(2,0,"","")
	else
	stats="�¼���¼�б�"
	call nav()
	call head_var(1,BoardDepth,0,0)
	end if
	if Cint(GroupSetting(39))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���������̳�¼���Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
		founderr=true
	end if

	'###################��������޸Ŀ�ʼ(asilas����)##################
	if boardid>0 then
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
	end if
	'####################��������޸Ľ���(asilas����)#################

	if master then founderr=false
	if founderr then
		call dvbbs_error()
	else
		if request("action")="dellog" then
			call batch()
		else
			call boardeven()
		end if
		if founderr then call dvbbs_error()
		call activeonline()
	end if
	call footer()
	REM ��ʾ������Ϣ---Headinfo
	sub boardeven()
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	if master then
	response.write "<div align=center>���������Ա��������ʱ���л�������״̬</div>"
	end if

	response.write "<form action=bbseven.asp?action=dellog&boardid="&boardid&" method=post name=even>"
	response.write "<input type=hidden name=boardid value="&boardid&">"
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableborder1>"
	response.write "<tr align=center>"
	response.write "<th width=15% height=25>����</td>"
	response.write "<th width=50% id=tabletitlelink>�¼�����(<a href=bbseven.asp>�鿴�����¼���¼</a>)</td>"
	response.write "<th width=20% id=tabletitlelink><a href=bbseven.asp?action=batch&boardid="&boardid&">����ʱ��</a></td>"
	response.write "<th width=15% >������</td></tr>"


	set rs=server.createobject("adodb.recordset")
	if boardid>0 then
	sql="select * from log where l_boardid="&boardid&" order by l_addtime desc"
	else
	sql="select * from log order by l_addtime desc"
	end if
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		response.write "<tr><td class=tablebody1 colspan=4 height=25>���滹û���κ��¼�</td></tr>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)

		if bgcolor=Forum_body(4) then
		bgcolor=Forum_body(5)
		else
		bgcolor=Forum_body(4)
		end if

		response.write "<tr>"
		response.write "<td class=tablebody1 align=center height=24><a href=dispuser.asp?name="&htmlencode(rs("l_touser"))&" target=_blank>"&htmlencode(rs("l_touser"))&"</a></td>"
		response.write "<td class=tablebody1>"&htmlencode(rs("l_content"))&"</td>"
		response.write "<td class=tablebody1>"
		if request("action")="batch" and (master or boardmaster) then
		response.write "<input type=checkbox name=lid value="&rs("l_id")&">"
		end if
		response.write rs("l_addtime")
		response.write "</td>"
		response.write "<td align=center class=tablebody1>"

		if master or superboardmaster then
			response.write "<a href=dispuser.asp?name="&htmlencode(rs("l_username"))&" target=_blank>"&htmlencode(rs("l_username"))&"</a>"
		elseif boardid=0 and not (master or superboardmaster) then
			response.write "����"
		elseif Board_Setting(36)<>"" and isnumeric(Board_Setting(36)) then
			if Cint(Board_Setting(36))=0  then
			response.write "<a href=dispuser.asp?name="&htmlencode(rs("l_username"))&" target=_blank>"&htmlencode(rs("l_username"))&"</a>"
			else
			response.write "����"
			end if
		else
			response.write "����"
		end if


		response.write "</td></tr>"
		page_count = page_count + 1
		rs.movenext
		wend
	end if
	if request("action")="batch" then
		response.write "<tr><td class=tablebody2 colspan=4>��ѡ��Ҫɾ�����¼���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=act value=ɾ��  onclick=""{if(confirm('��ȷ��ִ�д˲�����?')){this.document.even.submit();return true;}return false;}"">"&_
				"��<input type=submit name=act onclick=""{if(confirm('ȷ���������վ���еļ�¼��?')){this.document.even.submit();return true;}return false;}"" value=�����־></td></tr>"
	end if
	response.write "</table>"

  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"
	response.write "<tr><td valign=middle nowrap>"
	response.write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
	response.write "&nbsp;ÿҳ<b>"&Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"
	response.write "<td valign=middle nowrap align=right>��ҳ��"
	call DispPageNum(currentpage,PCount,"""?page=","&boardid="&boardid&"""")
	response.write "</td></tr></table>"
	rs.close
	set rs=nothing
	end sub

	sub batch()
	dim lid
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
	end if
	if boardid=0 then
		if not (master or  superboardmaster) then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>������ϵͳ����Ա�����ܹ���������־��"
		end if
	else
		if not boardmaster then
			Errmsg=Errmsg+"<br>"+"<li>�����Ǹð����������ϵͳ����Ա��<br><li>������û��ʹ�øù��ܵ�Ȩ�ޡ�"
			founderr=true
		end if
	end if
	if request("act")="ɾ��" then
		if request.form("lid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ������¼���"
		else
		lid=replace(request.Form("lid"),"'","")
		lid=replace(lid,";","")
		lid=replace(lid,"--","")
		lid=replace(lid,")","")
		end if
	end if

	if founderr then exit sub

	if request("act")="ɾ��" then
	conn.execute("delete from log where l_id in ("&lid&")")
	elseif request("act")="�����־" then
		if  boardmaster then
	conn.execute("delete from log where l_boardid="&boardid&" ")
		else
	conn.execute("delete from log  ")
		end if
	end if
	sucmsg="<li>ɾ��ָ���¼��ɹ�"
	call dvbbs_suc()
	end sub
	%>