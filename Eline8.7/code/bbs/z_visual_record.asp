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
	title="出售记录"
case "send"
	title="赠送记录"
case "recv"
	title="受赠记录"
case else
	title="购买记录"
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
	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
stats="查看"&title
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
			<th height=25 colspan=2><font color=<%=forum_body(8)%>><%=membername%></font>：形象商品--<%=title%></th>
		</tr>
		<%if rsvisual.eof or rsvisual.bof then%>
			<tr>
				<td height=25 colspan=2 align=center class=tablebody1><%
					select case curmenu
					case "send"
						response.write "没有赠送任何商品！"
					case "recv"
						response.write "没有受赠任何商品！"
					case "sale"
						response.write "没有出售任何商品！"
					case else
						response.write "没有购买任何商品！"
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
					<td class=tablebody1 height=20 width=370>&nbsp;<b>商品名称：</b><%
						response.write rsvisual("name")
						response.write "("
						if rsvisual("sex")=1 then
							response.write "男"
						elseif rsvisual("sex")=0 then
							response.write "女"
						else
							response.write "不限"
						end if
						response.write ")"
					%></td>
				</tr>
				<tr valign=middle>
					<td class=tablebody1 height=20 width=370>&nbsp;<b><%
						select case curmenu
						case "send"
							response.write "赠送价格："
						case "recv"
							response.write "受赠价格："
						case "sale"
							response.write "出售价格："
						case else
							response.write "购买价格："
						end select
						if rsvisual("price")=0 then
							response.write "<font color="&forum_body(8)&">免费</font>"
						else
							response.write rsvisual("price")&" 元"
						end if
					%></td>
				</tr>
				<tr valign=middle>
					<%select case curmenu
					case "send"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>受 赠 者：</b><a href=dispuser.asp?name=<%=rsvisual("username")%> target=_blank><font color=blue><%=rsvisual("username")%></font></a></td>
					<%case "recv"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>赠 送 者：</b><a href=dispuser.asp?name=<%=rsvisual("fromuser")%> target=_blank><font color=blue><%=rsvisual("fromuser")%></font></a></td>
					<%case "sale"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>购 买 者：</b><%
							if isnull(rsvisual("username")) or rsvisual("username")="" then
								response.write "<font color="&forum_body(8)&">尚未售出</font>"
							else
								response.write "<a href=dispuser.asp?name="&rsvisual("username")&" target=_blank><font color=blue>"&rsvisual("username")&"</font></a>"
							end if
						%></td>
					<%case else%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>卖 出 者：</b><%
							if rsvisual("fromuser")="#形象商店#" then
								response.write "<B><font color=blue>形象商店</font></B>"
							else
								%><a href=dispuser.asp?name=<%=rsvisual("fromuser")%> target=_blank><font color=blue><%=rsvisual("fromuser")%></font></a><%
							end if
						%></td>
					<%end select%>
				</tr>
				<tr valign=middle>
					<%select case curmenu
					case "send"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>赠送时间：</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case "recv"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>受赠时间：</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case "sale"%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>出售时间：</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
					<%case else%>
						<td class=tablebody1 height=20 width=370>&nbsp;<b>购买时间：</b><%=formatdatetime(rsvisual("dateandtime"),2)%></td>
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