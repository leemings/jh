<!--#include file="conn.asp"-->
<!--#include file="z_dgconn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"--> 
<%
'---------------------
' write by ��ˮ��ɽ
'---------------------
	stats="�ҵ�ף���б�"
	
	dim abgcolor,username

	call nav()
	call head_var(0,0,"���������","z_dglistall.asp")
	
	if not founduser then
		errmsg=errmsg+"<br>"+"<li>���˲����Բ��ĵ���б�"
		errmsg=errmsg+"<br>"+"<li>����<a href=login.asp><font color=blue>��¼</font></a>,��û��<a href=reg.asp><font color=blue>ע��</font></a>"
		founderr=true
	end if
	
	if founderr then 
		call dvbbs_error()
	else

		dim page_count,Pcount,totalrec,mytotalrec,totalPages,currentPage	
		
		set rs=server.createobject("adodb.recordset")
		set rs=connDG.execute("select count(sender) from media")
			totalrec=rs(0)
		rs.close	
		set rs=connDG.execute("select count(sender) from media where incept='"&membername&"' or incept='ȫ���Ա'")
			mytotalrec=rs(0)
		rs.close		
	%>	
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
			<tr>
				<th><a href=z_dglistall.asp><font color=white><b>���е���б�</b></font></a></th>
				<th><a href=z_dglistme.asp><font color=white><b>�ҵĵ���б�</b></font></a></th>
				<th><a href=z_dgwrite.asp><font color=white><b>��Ҫ���</b></font></a></th>
			</tr>
			<tr>
				<td class=tablebody2 align=center valign=middle colspan="3">���ף���嵥����[<font color=red><b><%=mytotalrec%></b></font>]����<a href=z_dglistall.asp><b>��̳�ܵ���б�</b></a>�嵥����[<font color=red><b><%=totalrec%></b></font>]����<font color=green><ֱ�ӵ����������></font></td>
			</tr>			
		</table>
		<br>
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
			<tr>
				<th width=80>�����</th><th width=80>�Է�����</th><th width=220>����</th><th width=120>ʱ��</th><th width=*>ף����</th><th width=50>����</th> 
			</tr>	
	<%
		sql="select * from media where incept='"&membername&"' or incept='ȫ���Ա' Order By sendtime Desc"
		rs.open sql,connDG,3,3				
		if rs.eof and rs.bof then
			currentpage=0
	%>		
			<tr><td class=tablebody2 valign=middle colspan=6>��ǰû�е���б�</td></tr>
	<%		
		else

			currentPage=request.querystring("page")
			if currentpage="" or not isInteger(currentpage) then
				currentpage=1
			else
				currentpage=clng(currentpage)
				if err then
					currentpage=1
					err.clear
				end if
			end if

		rs.PageSize = 10
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)
%>
				<tr>
					<td class=tablebody2 align=center valign=middle><a href=javascript:openScript('dispuser.asp?name=<%=rs("sender")%>',650,500)><font color=blue><%=rs("sender")%></font></a></td>
					<td class=tablebody1 align=center valign=middle><%if trim(rs("incept"))<>"ȫ���Ա" then%><a href=javascript:openScript('dispuser.asp?name=<%=rs("incept")%>',650,500)><font color=green><%=rs("incept")%></font></a><%else%><font color=olive><%=rs("incept")%></font><%end if%></td>
					<td class=tablebody2 align=center valign=middle><a href=z_dgplay.asp?url=<%=replace(rs("url"),chr(32),"%20",1)%>&medianame=<%=replace(rs("medianame"),chr(32),"%20")%> target=_blank><%=rs("medianame")%></a></td>
					<td class=tablebody1 align=center valign=middle><%=rs("sendtime")%></td>
					<td class=tablebody1 align=left valign=middle><font color=red><%=DvBCode(rs("content"),1,2)%></font></td>
					<td class=tablebody2 align=center valign=middle><a href=z_dgdel.asp?id=<%=rs("id")%> onclick="javascript:{if(confirm('��ȷ��ִ��ɾ��������?')){return true;}return false;}">[ɾ��]</a></td>					
				</tr>
<%
	  	page_count = page_count + 1
		rs.movenext
		wend
	end if
%>
</table>
<%
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			" ÿҳ<b>10</b>�� ����<b>"&mytotalrec&"</b>�����</td>"&_
			"<td valign=middle nowrap><div align=right><p>��ҳ�� "
	call DispPageNum(currentpage,PCount,"""?page=","""")
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing
'=========================
%>
		<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
			<tr>
				<th><font color=white>--== ��̳���̨-�ҵ�ף���б� ==--</font></th>
			</tr>
		</table>
<%		
	end if
	call CloseDB()
	call activeonline()
	call footer()
%>