<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	dim title,currentPage,idlist
 	const MaxPerPage=25
 	title=request("txtitle")
 	if not isempty(request("page")) then
		currentPage=cint(request("page"))
 	else
		currentPage=1
 	end if
	if not isempty(request("delAnnounce")) then
		idlist=request("delAnnounce")
		if instr(idlist,",")>0 then
			dim idarr
			idArr=split(idlist)
			dim id
			for i = 0 to ubound(idarr)
				id=clng(idarr(i))
				if request("agree")="��׼" then
					call agreeannounce(clng(id))
					response.write "�� ׼ �� ����<br>"
					response.end
				end if
				if request("noagree")="����׼" then
					call noagreeannounce(clng(id))
					response.write " �� �� �� ����<br>"
					response.end
				end if
			next
		else
			if request("agree")="��׼" then
				call agreeannounce(clng(idlist))
				response.write "�� ׼ �� ����<br>"
				response.end
			end if
			if request("noagree")="����׼" then
				call noagreeannounce(clng(idlist))
				response.write " �� �� �� ����<br>"
				response.end
			end if
		end if
	end if%>
	<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td width="100%" align="center"><form method=Post action="?"><%
				Set rs= Server.CreateObject("ADODB.Recordset")
				sql="select * from [user] where vip=4 order by userid desc"
				rs.open sql,conn,1,1
				dim totalPut,pagecount,curcount
			 	if rs.eof and rs.bof then
			 		response.write "�� û �� �� �� �� ��"
			 	else
			 		totalPut=rs.recordcount
					if totalPut mod MaxPerPage=0 then
						pagecount=totalPut \ MaxPerPage
					else
						pagecount=totalPut \ MaxPerPage+1
					end if
					if currentpage > pagecount then currentpage = pagecount
					if currentpage < 1 then currentpage=1
					rs.PageSize = MaxPerPage
					rs.AbsolutePage=currentpage
					curcount=0
					%><table class="tableBorder" cellspacing="1" width="95%" cellpadding="0" align="center">
 						<tr>
 							<th width="30%" align="center" height="20" >�û���</td>
							<th width="30%" align="center">ע��ʱ��</td>
							<th width="40%" align="center"><input type='submit' name="agree" class=buttonface value='��׼'>&nbsp; <input type='submit' name="noagree" class=buttonface value='����׼'>ȫѡ��<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)"> </th>
						</tr>
						<%do while not rs.eof
							curcount=curcount+1
							%><tr>
								<td height="23" width="30%" class=forumrow align="center"><a href=DISPUSER.ASP?name=<%=rs("username")%> target=_blank><%=rs("username")%></a></td>
								<td width="30%" class=forumrow align="center"><%=rs("addDate")%></td>
								<td width="40%" class=forumrow align="center"><input type='checkbox' name='delAnnounce' value='<%=cstr(rs("UserID"))%>'></td>
							</tr>
							<%rs.movenext
						loop%>
						<tr>
							<td height="23" colspan=3 align="right" class=forumrow>��ҳ��<%call DispPageNum(currentpage,pagecount,"?page=","")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
					</table>
				<%end if
			%></td>
		</tr>
	</form></table>
<%end if%>
</body>

<script language="javascript">
<!--
function CheckAll(form) {
 for (var i=0;i<form.elements.length;i++) {
 var e = form.elements[i];
 if (e.name != 'chkall') e.checked = form.chkall.checked; 
 }
 }
//-->
</script>

<%sub noagreeannounce(id)
	dim rss,rs2,sql2,rs,sql
	set rss=conn.execute("select * from [user] where userid="&cstr(id)&"")
	set rs2=server.createobject("adodb.recordset") 
	sql2="select * from [message]" 
	rs2.open sql2,conn,1,3 
	rs2.addnew 
	rs2("sender")=forum_info(0) 
	rs2("incept")=rss("username") 
	rs2("title")="����VIP��Ա����û�еõ���׼��" 
	rs2("content")="����VIP��Ա����û�еõ���׼���������Ա��ϵ��" 
	rs2("flag")=0 
	rs2("issend")=1 
	rs2.update 
	rs2.close 
	rss.close
	conn.execute("update [user] set vip=0 where userid="&cstr(id)&"")
End sub
 
sub agreeannounce(id)
	dim rs2,sql2,user
	user=conn.execute("select username from [user] where userid="&cstr(id)&"")(0)
	conn.execute("update [user] set vip=1,vipdate=now(),userwealth2=userwealth,userep2=userep,usercp2=usercp,userpower2=userpower,article2=article where userid="&cstr(id))
	set rs2=server.createobject("adodb.recordset") 
	sql2="select * from [message]" 
	rs2.open sql2,conn,1,3 
	rs2.addnew 
	rs2("sender")=forum_info(0) 
	rs2("incept")=user 
	rs2("title")="����VIP��Ա�����Ѿ��õ���׼��" 
	rs2("content")="����VIP��Ա�����Ѿ��õ���׼���������Ѿ�����ʽ��VIP��Ա�ˣ�" 
	rs2("flag")=0 
	rs2("issend")=1 
	rs2.update 
	rs2.close
End sub%>
