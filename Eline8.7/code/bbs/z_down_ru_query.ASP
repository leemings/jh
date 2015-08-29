<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<%dim MaxPerPage
stats="软件搜索"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
MaxPerPage=sonum	'每页显示程序最大数量
dim totalPut 
dim CurrentPage
dim n
dim j
dim keyword
dim findword

keyword=trim(checkstr(request("keyword")))
keyword=replace(replace(replace(keyword,"<","&lt;"),">","&gt;")," ","&nbsp;")
if keyword="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入查询条件。"
	call dvbbs_error()
else
	if request("action")="title" then
		findword="showname like '%"&keyword&"%' "
	else
		findword="note like '%"&keyword&"%' "
	end if
 	if not isempty(request("page")) then
 		currentPage=cint(request("page"))
 	else
 		currentPage=1
 	end if
 	set rs=server.createobject("adodb.recordset")
	dim classid,Nclassid
	dim classname,Nclassname
	if request("classid")="" then
		classid=""
		classname="所有软件"
	else
		classid="classid="&cstr(request("classid"))&" and "
		sql="select class from Dclass where classid="&cstr(request("classid"))
		rs.open sql,conndown,1,1
		classname=rs("class")
		rs.close
	end if
	if request("Nclassid")="" then
		Nclassid=""
		Nclassname="所有软件"
	else
		Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and "
		sql="select Nclass from DNclass where Nclassid="&cstr(request("Nclassid"))
		rs.open sql,conndown,1,1
		Nclassname=rs("Nclass")
		rs.close
	end if%>
	<!-- #include file="z_down_forum_js.asp" -->
	<%call down_nav()%>
	<table align=center border="0" width="<%=Forum_body(12)%>" cellspacing=0 cellpadding=0>
		<tr>
			<td width="100%" valign="top" align="right">
				<table border="0" cellspacing="1" cellpadding="3" align=center class=tableborder1>
					<tr>
						<td align=center height=25 colspan=5 class=tablebody2><%
							response.write "查询条件：<font color=red><b>"&keyword&"</b></font>"
							response.write "在"&classname&"的"&Nclassname&"中"
						%></td>
					</tr>
					<tr>
						<th width="41%" height=25>软件名称和简介</th>
						<th width="11%">推荐</th>
						<th width="20%">更新日期</th>
						<th width="14%">下载次数</th>
						<th width="14%">文件大小</th>
					</tr>
					<%set rs=server.createobject("adodb.recordset")
					if request("Nclassid")="" then
						sql="select * from download where "&classid&" "&findword&" "
						sql=sql&" order by id desc"
					else
						sql="select * from download where "&classid&" "&Nclassid&" "&findword&" "
						sql=sql&" order by id desc"
					end if
					rs.open sql,conndown,1,1
					if rs.eof and rs.bof then 
						response.write "<tr><td colspan=5 align=center height=25 class=tablebody1>没有或尚未找到任何程序</td></tr>" 
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
								<td width="100%" colspan="5" height=25 class=tablebody1>&nbsp;<a href="z_down_list.asp?id=<%=rs("id")%>"><%=replace(rs("showname"),""&keyword&"","<font color=red>"&keyword&"</font>")%></a></td>
							</tr>
							<tr>
								<td width="41%" height=25 class=tablebody2>&nbsp;运行环境：<%=rs("system")%> 授权方式：<%=rs("orders")%></td>
								<td width="11%" align=center class=tablebody2><%if rs("hot")>3 then%><font color=red>推荐</font><%end if%></td>
								<td width="20%" align=center class=tablebody2><%=rs("dateandtime")%></td>
								<td width="14%" align=center class=tablebody2><%=rs("hits")%></td>
								<td width="14%" align=center class=tablebody2><%=rs("size")%></td>
							</tr>
							<tr>
								<td width="100%" colspan="5" height=25 class=tablebody1>&nbsp;<%if len(rs("note"))>50 then%><%=left(rs("note"),50)%>……<%else%><%=rs("note")%><%end if%></td>
							</tr>
							<%rs.movenext
						loop%>
						<form method=Post action=?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=<%=request("order")%>&updown=<%=request("updown")%>>
						<tr>
							<td colspan=2 height=25 class=tablebody2>&nbsp;<font color=<%=forum_body(8)%>><b><%=Nclassname%></b></font>&nbsp;&nbsp;页次：<b><%=currentPage%></b> / <b><%=n%></b> 页 每页 <b><%=maxperpage%></b> 软件数 <b><%=totalPut%></b></td>
							<td colspan=3 class=tablebody2 align=right>分页：<%call disppagenum(CurrentPage,n,"?page=","&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&action="&request("action")&"&keyword="&keyword)%>&nbsp;转到：<input type=text name=Page size=3 maxlength=10  value=<%=currentpage%>><input type=submit value=Go name=submit>&nbsp;</td>
						</tr>
						</form>
					<%end if
			 		rs.close%>
			 	</table>
				<!--#include file=z_down_query.asp-->
			</td>
		</tr>
	</table>
	<%call down_payuser()
	call down_manage()
end if
conndown.close
set conndown=nothing
call activeonline()
call footer()%>
 








