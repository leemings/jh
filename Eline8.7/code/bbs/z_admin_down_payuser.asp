<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file="z_down_function.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim MaxPerPage,title,currentpage,idlist,rs1,mess
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	MaxPerPage=eventnum
	title=request("txtitle")
	if not isempty(request("page")) then
		currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	if not isempty(request("delAnnounce")) then
		idlist=request("delAnnounce")
		set rs=conndown.execute("select * from [user] where id in ("&idlist&")")
		do while not rs.eof
			conndown.execute("delete from [userevent] where username='"&rs("username")&"'")
			conndown.execute("delete from [downevent] where username='"&rs("username")&"'")
			rs.movenext
		loop
		rs.close
    conndown.execute("delete from [user] where id in ("&idlist&")")
		response.write " ɾ �� �� ����"
		response.end
	end if 
	if not isempty(request("lockAnnounce")) then
		idlist=request("lockAnnounce")
		set rs=conndown.execute("select * from [user] where id in ("&idlist&")")
		do while not rs.eof
			if rs("lockuser")=1 then
				conndown.execute("update [user] set lockuser=2 where id="&rs("id"))
				set rs1=server.createobject("adodb.recordset")
				sql="select * from events"
				rs1.open sql,conndown,1,3
				mess=""&membername&"�����˸����û���"&rs("username")&"��"	
				rs1.addnew
				rs1("event")=mess
				rs1("addtime")=date()
				rs1.update
				rs1.close
				sql="select * from [message]"      
				rs1.open sql,conn,1,3      
				rs1.addnew      
				rs1("sender")=forum_info(0)&"��������"      
				rs1("incept")=rs("username")
				rs1("title")="���ĸ��������û��˺��Ѿ���������"      
				rs1("content")="���ĸ��������û��˺��Ѿ����������������Ա��ϵ��"      
				rs1("flag")=0      
				rs1("issend")=1      
				rs1.update      
				rs1.close
				sql="select * from [userevent]"
				rs1.open sql,conndown,1,3
				rs1.addnew
				rs1("username")=rs("username")
				rs1("userevent")="���ĸ��������û��˺ű�������"
				rs1("addtime")=date()
				rs1("conuser")=membername
				rs1.update
				rs1.close
				set rs1=nothing
			else  	
				conndown.execute("update [user] set lockuser=1 where id="&rs("id"))
				set rs1=server.createobject("adodb.recordset")
				sql="select * from events"
				rs1.open sql,conndown,1,3
				mess=""&membername&"�����˸����û���"&rs("username")&"��"	
				rs1.addnew
				rs1("event")=mess
				rs1("addtime")=date()
				rs1.update
				rs1.close
				sql="select * from [message]"      
				rs1.open sql,conn,1,3      
				rs1.addnew      
				rs1("sender")=forum_info(0)&"��������"      
				rs1("incept")=rs("username")
				rs1("title")="���ĸ��������û��˺��Ѿ���������"      
				rs1("content")="���ĸ��������û��˺��Ѿ���������������ʹ�ã�"      
				rs1("flag")=0      
				rs1("issend")=1      
				rs1.update      
				rs1.close
				sql="select * from [userevent]"
				rs1.open sql,conndown,1,3
				rs1.addnew
				rs1("username")=rs("username")
				rs1("userevent")="���ĸ��������û��˺��ѱ�������"
				rs1("addtime")=date()
				rs1("conuser")=membername
				rs1.update
				rs1.close
				set rs1=nothing
			end if
			rs.movenext
		loop
		rs.close
		response.write "�� �� �� ����"
		response.end
	end if%>
	<table border="0" width="95%" cellspacing="0" cellpadding="0" align=center>
		<form method=Post action="?">
		<tr>
			<td width="100%"><font color=<%=forum_body(8)%>>�����û�����</font> | <a href=z_admin_down_addpayuser.asp>���Ӹ����û�</a></td>
		</tr>
		<tr>
			<td width="100%" align=center>
				<table border="0" cellspacing="1" cellpadding="3" width="100%" class=tableborder>
					<tr>
					  <th width="20%" height="25">�û���</th>
						<th width="20%">ע��ʱ��</th>
						<th width="20%">��������</th>
						<th width="20%">״̬</th>
						<th width="10%"><input type='submit' value='�� ��'></th>
						<th width="10%"><input type='submit' value='ɾ ��'></th>
					</tr>
					<%dim n,totalput
					sql="select * from [user] where apply=1 order by id desc"
					Set rs= Server.CreateObject("ADODB.Recordset")
					rs.open sql,conndown,1,1
					if rs.eof and rs.bof then
						response.write "<tr height=25><tr class=forumrow align=center colspan=6>�� û �� �� �� �� ��</td></tr>"
					else
						totalPut=rs.recordcount
						if totalPut mod MaxPerPage=0 then
							n= totalPut \ MaxPerPage
						else
							n= totalPut \ MaxPerPage+1
						end if
						if currentpage > n then currentpage = n
						if currentpage < 1 then currentpage = 1
						rs.PageSize = MaxPerPage
						rs.AbsolutePage=currentpage
						i=0
						do while not rs.eof and i<MaxPerPage
							i=i+1%>
							<tr>
								<td height="23" align="center" class=forumrow><%=rs("username")%></td>
								<td align="center" class=forumrow><%=rs("regtime")%></td>
								<td align="center" class=forumrow><%=rs("allpoint")%></td>
								<td align="center" class=forumrow><%if rs("lockuser")=1 then%>����<%else%>����<%end if%></td>
								<td align="center" class=forumrow><input type='checkbox' name='lockAnnounce' value='<%=cstr(rs("ID"))%>'></td>
								<td align="center" class=forumrow><input type='checkbox' name='delAnnounce' value='<%=cstr(rs("ID"))%>'></td>
							</tr>
							<%rs.movenext
						loop%>
						<tr>
							<td height="25" class=forumrow align=right colspan=6>ҳ�Σ�<b><%=currentPage%></b> / <b><%=n%></b> ҳ ÿҳ <b><%=maxperpage%></b> �¼��� <b><%=totalPut%></b>&nbsp;&nbsp;&nbsp;&nbsp;��ҳ��<%call disppagenum(currentpage,n,"?page=","")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
					<%end if
					rs.close
					set rs=nothing%>
				</table>
			</td>
		</tr>
		</form>
	</table>
<%end if%>
</body>
