<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
<tr>
<th align=left height=25><b>-=> <a href=z_dgwrite.asp>���ף����</a></b></th>
</tr>
<tr>
<td class=tablebody1 height=80 width="100%">
	<%dim page_count,Pcount,totalrec,mytotalrec,totalPages,currentPage
	dim abgcolor,username	
	set rs=server.createobject("adodb.recordset")
	sql="select * from media Order By sendtime Desc"
	rs.open sql,connDG,3,3				
	if rs.eof and rs.bof then
		%>��ǰû�е���б�<%		
	else
		%><marquee scrollamount=1 direction=up height=80 onmouseover=this.stop(); onmouseout=this.start();><%
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
			if trim(rs("incept"))=membername or trim(rs("incept"))="ȫ���Ա" then
				%><a href=javascript:openScript('dispuser.asp?name=<%=rs("sender")%>',650,500)><font color=blue><%=rs("sender")%></font></a> �� <%=rs("sendtime")%> Ϊ <%if trim(rs("incept"))=membername then%><a href=javascript:openScript('dispuser.asp?name=<%=rs("incept")%>',650,500)><font color=#CC66FF><%=rs("incept")%></font></a><%else%><font color=olive><%=rs("incept")%></font><%end if%> ����һ�� <a href='z_dgplay.asp?url=<%=replace(rs("url"),chr(32),"%20",1)%>&medianame=<%=replace(rs("medianame"),chr(32),"%20")%>' target=_blank><font color=red>[<%=rs("medianame")%>]</font></a>	��ף������<font color=orange><B><%=DvBCode(rs("content"),1,2)%></B></font><br><%		
			else
				%><a href=javascript:openScript('dispuser.asp?name=<%=rs("sender")%>',650,500)><font color=blue><%=rs("sender")%></font></a> �� <%=rs("sendtime")%> Ϊ <a href=javascript:openScript('dispuser.asp?name=<%=rs("incept")%>',650,500)><font color=green><%=rs("incept")%></font></a> ����һ�� <a href='z_dgplay.asp?url=<%=replace(rs("url"),chr(32),"%20",1)%>&medianame=<%=replace(rs("medianame"),chr(32),"%20")%>' target=_blank><font color=red>[<%=rs("medianame")%>]</font></a> ��ף������<font color=orange>******�Ѿ�����******</font><br><%
			end if
	  	page_count = page_count + 1
			rs.movenext
		wend
	end if
%></table>
<BR>
