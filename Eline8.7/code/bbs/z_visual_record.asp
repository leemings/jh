<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp"-->
<%dim menu,title,curpage
if not isempty(request("page")) then
	curPage=cint(request("page"))
else
	curPage=1
end if
menu=request("menu")
select case menu
case "sale"
	title="���ۼ�¼"
case "send"
	title="���ͼ�¼"
case "recv"
	title="������¼"
case else
	title="�����¼"
end select%>
<html>
<head>
<title><%=Forum_info(0)%>--<%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>
<%
if not founduser then
	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
stats="�鿴"&title
if founderr then
	call dvbbs_error()
else
	call list(menu)
end if
call activeonline()
call footer()

sub list(curmenu)
	dim rsvisual,totalrec,n
	set rsvisual=server.createobject("ADODB.Recordset")
	select case curmenu
	case "send"
		sql="select e.dateandtime,e.price,e.username,e.fromuser,i.name,i.content,i.sex,i.id from visual_events e inner join visual_items i on e.itemid=i.id where e.fromuser='"&membername&"' and e.type=1 order by e.dateandtime desc"
	case "recv"
		sql="select e.dateandtime,e.price,e.username,e.fromuser,i.name,i.content,i.sex,i.id from visual_events e inner join visual_items i on e.itemid=i.id where e.username='"&membername&"' and e.type=1 order by e.dateandtime desc"
	case "sale"
		sql="select e.dateandtime,e.price,e.username,e.fromuser,i.name,i.content,i.sex,i.id from visual_events e inner join visual_items i on e.itemid=i.id where e.fromuser='"&membername&"' and e.type=2 order by e.dateandtime desc"
	case else
		sql="select e.dateandtime,e.price,e.username,e.fromuser,i.name,i.content,i.sex,i.id from visual_events e inner join visual_items i on e.itemid=i.id where e.username='"&membername&"' and e.type=0 order by e.dateandtime desc"
	end select
	rsvisual.open sql,conn,1,1
	%>
	<table width=95% cellpadding=3 cellspacing=1 border=0 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr>
			<th height=25 colspan=2><font color=<%=forum_body(8)%>><%=membername%></font>��������Ʒ--<%=title%></th>
		</tr>
		<%if rsvisual.eof or rsvisual.bof then%>
			<tr>
				<td height=25 colspan=2 align=center class=tablebody1><%
					select case curmenu
					case "send"
						response.write "û�������κ���Ʒ��"
					case "recv"
						response.write "û�������κ���Ʒ��"
					case "sale"
						response.write "û�г����κ���Ʒ��"
					case else
						response.write "û�й����κ���Ʒ��"
					end select
				%></td>
			</tr>
		<%else
			totalrec=rsvisual.recordcount
			if totalrec mod 3=0 then
				n= totalrec \ 3
			else
				n= totalrec \ 3+1
			end if
			if curpage > n then curpage = n
			if curpage < 1 then curpage = 1
			rsvisual.PageSize = 3
			rsvisual.AbsolutePage=curpage
			i=0
			do while not rsvisual.eof and i<3
				i=i+1%>
				<tr valign=middle>
					<td align=center class=tablebody1 rowspan=4 baseName="images/img_visual/blank.gif" width="84" height="84" border="0" itemgender="" itemno="<%=rsvisual("id")%>" layerno="" localpic="<%=LocalPic%>" ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
					<td class=tablebody1 height=20 width=370>&nbsp;<b>��Ʒ���ƣ�</b><%
						response.write rsvisual("name")
						response.write "("
						if rsvisual("sex")=1 then
							response.write "��"
						elseif rsvisual("sex")=0 then
							response.write "Ů"
						else
							response.write "����"
						end if
						response.write ")"
					%></td>
				</tr>
				<tr valign=middle>
					<td class=tablebody1 height=20 width=370>&nbsp;<b><%
						select case curmenu
						case "send"
							response.write "���ͼ۸�"
						case "recv"
							response.write "�����۸�"
						case "sale"
							response.write "���ۼ۸�"
						case else
							response.write "����۸�"
						end select
						if rsvisual("price")=0 then
							response.write "<font color="&forum_body(8)&">���</font>"
						else
							response.write rsvisual("price")&" Ԫ"
						end if
					%></td>
				</tr>
				<tr valign=middle>
					<%select case curmenu
					case "send"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>�� �� �ߣ�</b><a href=dispuser.asp?name=<%=rsvisual("username")%> target=_blank><font color=blue><%=rsvisual("username")%></font></a></td>
					<%case "recv"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>�� �� �ߣ�</b><a href=dispuser.asp?name=<%=rsvisual("fromuser")%> target=_blank><font color=blue><%=rsvisual("fromuser")%></font></a></td>
					<%case "sale"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>�� �� �ߣ�</b><%
							if isnull(rsvisual("username")) or rsvisual("username")="" then
								response.write "<font color="&forum_body(8)&">��δ�۳�</font>"
							else
								response.write "<a href=dispuser.asp?name="&rsvisual("username")&" target=_blank><font color=blue>"&rsvisual("username")&"</font></a>"
							end if
						%></td>
					<%case else%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>�� �� �ߣ�</b><%
							if rsvisual("fromuser")="#�����̵�#" then
								response.write "<B><font color=blue>�����̵�</font></B>"
							else
								%><a href=dispuser.asp?name=<%=rsvisual("fromuser")%> target=_blank><font color=blue><%=rsvisual("fromuser")%></font></a><%
							end if
						%></td>
					<%end select%>
				</tr>
				<tr valign=middle>
					<%select case curmenu
					case "send"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>����ʱ�䣺</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case "recv"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>����ʱ�䣺</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case "sale"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>����ʱ�䣺</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case else%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>����ʱ�䣺</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%end select%>
				</tr>
				<%rsvisual.movenext
			loop%>
			<tr>
				<td class=tablebody2 height=25 align=center colspan=2 width=100%><%call disppagenum(curpage,n,"?page=","&menu="&curmenu)%></td>
			</tr>
		<%end if%>
	</table>
	<%rsvisual.close
	set rsvisual=nothing
end sub
%>