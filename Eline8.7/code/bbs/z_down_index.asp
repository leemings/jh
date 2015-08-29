<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<!--#include file="z_down_function.asp" -->
<!--#include file="inc/ubbcode.asp" -->
<%stats="软件列表"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")

dim MaxPerPage	'每页显示程序最大数量
dim totalPut   
dim CurrentPage,n
dim TotalPages
dim j
dim keyword
dim updown
dim order_name

MaxPerPage=listnum
order_name=Request("Order")
if not isempty(request("page")) then
	currentPage=cint(request("page"))
else
	currentPage=1
end if
if request("updown")<>"" then
	updown="desc"
else
	updown=""
end if
select case order_name
case "showname"
	order_name="showname"
case "hot"
	order_name="hot"
case "dateandtime"
	order_name="dateandtime"
case "hits"
	order_name="hits"
case "orders"
	order_name="orders"
case "size"
	order_name="size"
case else
	order_name="id"
	updown="desc"
end select

dim classid,Nclassid
dim classname,Nclassname

set rs=server.createobject("adodb.recordset")
if request("classid")="" then
	classid="classid=1 and  "
	sql="select class from Dclass where classid=1"
	rs.open sql,conndown,1,1
	if rs.bof and rs.eof then
		response.write "还没有任何栏目，请到管理页面添加"
		response.end
	else
		classname=rs("class")
	end if
	rs.close
else
	classid="classid="&cstr(request("classid"))&" and  "
	sql="select class from Dclass where classid="&cstr(request("classid"))
	rs.open sql,conndown,1,1
	classname=rs("class")
	rs.close
end if
if request("Nclassid")="" then
	Nclassid=""
	Nclassname="所有软件"
else
	Nclassid=" Nclassid="&cstr(request("Nclassid"))&" and  "
	sql="select DNclass.Nclass,Dclass.class from DNclass,Dclass where DNclass.classid=Dclass.classid and DNclass.Nclassid="&cstr(request("Nclassid"))
	rs.open sql,conndown,1,1
	classname=rs("class")
	Nclassname=rs("Nclass")
	rs.close
end if%>
<!-- #include file="z_down_forum_js.asp" -->
<%call down_nav()%>
<table align=center border="0" width="<%=Forum_body(12)%>">
	<tr>
		<td width="24%" valign="top" align="left"><%
			if showset(5)=1 then
				%><!--#include file="z_down_allhot.asp"--><br><%
			end if
			if showset(6)=1 then
				%><!--#include file="z_down_hot.asp"--><br><%
			end if
			if showset(7)=1 then
				%><!--#include file="z_down_day.asp"--><br><%
			end if
			if showset(8)=1 then
				%><!--#include file="z_down_week.asp"--><br><%
			end if
			if showset(9)=1 then
				%><!--#include file="z_down_all.asp"--><br><%
			end if
		%></td>
		<td width="*" valign="top" align="right">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="50%"><a href=z_down_default.asp>下载分类</a>&nbsp;&gt;&gt;&nbsp;<a href='?classid=<%=request("classid")%>'><%=classname%></a>&nbsp;&gt;&gt;&nbsp;<%=Nclassname%><br><br></td>
					<td width="*" align=right nowrap><form action="?" method=get><select name="classid" size="1" onchange='javascript:submit()'><%
						set rs=server.createobject("adodb.recordset")
						sql="select class,classid from Dclass"
						rs.open sql,conndown,1,1
						if rs.bof and rs.eof then
							%><option value="">没有记录</option><%
						else
							do while not rs.eof
								%><option <%if request("classid")<>"" then%><%if cint(rs("classid"))=cint(request("classid")) then%> selected  <%end if%><%end if%>value="<%=rs("classid")%>"><%=rs("class")%></option><%
								rs.movenext
							loop
						end if
						rs.close
						%></select>&nbsp;<select name="Nclassid" size="1" onchange='javascript:submit()'><%
						if request("classid")="" then
							set rs=server.createobject("adodb.recordset")
							sql="select Nclass,Nclassid from DNclass"
						else
							sql="select Nclass,Nclassid from DNclass where classid="&request("classid")
						end if
						rs.open sql,conndown,1,1
						if rs.bof and rs.eof then
							%><option value="">没有记录</option><%
						else
							do while not rs.eof
								%><option <%if request("Nclassid")<>"" then%><%if cint(rs("Nclassid"))=cint(request("Nclassid")) then%> selected  <%end if%><%end if%>value="<%=rs("Nclassid")%>"><%=rs("Nclass")%></option><%
								rs.movenext
							loop
						end if
						rs.close
						%></select>&nbsp;<select name="order" size="1" onchange='javascript:submit()'>
						<option <%if request("order")<>"" then%><%if request("order")="id" then%> selected <%end if%><%end if%>value="id">按序号排序</option>
						<option <%if request("order")<>"" then%><%if request("order")="showname" then%> selected <%end if%><%end if%>value="showname">按标题排序</option>
						<option <%if request("order")<>"" then%><%if request("order")="hot" then%> selected <%end if%><%end if%>value="hot">按推荐排序</option>
						<option <%if request("order")<>"" then%><%if request("order")="hits" then%> selected <%end if%><%end if%>value="hits">按下载排序</option>
						<option <%if request("order")<>"" then%><%if request("order")="dateandtime" then%> selected <%end if%><%end if%>value="dateandtime">按日期排序</option>
					</select><input type=hidden name=updown value="desc"></form></td>
				</tr>
			</table>
			<table border="0" width="100%" cellspacing="1" cellpadding="3" align=center bgcolor=<%=Forum_Body(27)%>>
				<%if request("updown")="" then%>
					<tr>
						<th width="41%" height=22><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=showname&updown=desc" title="点击按升序排列"><font color="ffffff">软件名称和简介</font></a></th>
						<th width="11%"><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hot&updown=desc" title="点击按升序排列"><font color="ffffff">推 荐</font></a></th>
						<th width="20%"><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=dateandtime&updown=desc" title="点击按升序排列"><font color="ffffff">更新日期</font></a></th>
						<th width="14%"><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=orders&updown=desc" title="点击按升序排列"><font color="ffffff">授权形式</font></a></th>
						<th width="14%">文件大小</th>
					</tr>
				<%else%>
					<tr>
						<th width="41%" height=22><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=showname" title="点击按降序排列"><font color="ffffff">软件名称和简介</font></a></th>
						<th width="11%"><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=hot" title="点击按降序排列"><font color="ffffff">推 荐</font></a></th>
						<th width="20%"<a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=dateandtime" title="点击按降序排列"><font color="ffffff">更新日期</font></a></th>
						<th width="14%"><a href="?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=orders" title="点击按降序排列"><font color="ffffff">授权形式</font></a></th>
						<th width="14%">文件大小</th>
					</tr>
				<%end if
				set rs=server.createobject("adodb.recordset")
				if request("Nclassid")="" then
					sql="select * from download where "&classid&" stop=0"
					sql=sql&" order by "&order_name&" "&updown
				else
					sql="select * from download where "&classid&" "&Nclassid&" stop=0"
					sql=sql&" order by "&order_name&" "&updown
				end if
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then 
					response.write "<tr height=25><td colspan=5 align=center class=tablebody1>没有或未找到任何程序</td></tr>" 
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
							<td width="41%" height=22 class=tablebody2>&nbsp;<a href="z_down_list.asp?id=<%=rs("id")%>"><%=rs("showname")%></a></td>
							<td width="11%" align=center class=tablebody2><%if rs("hot")>3 then%><font color=red>推 荐</font><%end if%></td>    
							<td width="20%" align=center class=tablebody2><%=rs("dateandtime")%></td>
							<td width="14%" align=center class=tablebody2><%=rs("orders")%><br><%=downshow(rs("downshow"))%></td>
							<td width="14%" align=center class=tablebody2><%=rs("size")%></td>
						</tr>
						<tr>
							<td height=22 class=tablebody1 align=center>下载有效期：<%=downtime(rs("dateandtime"),rs("usetime"))%></td>
							<td colspan=2 class=tablebody1 align=center>本日下载：<%
								if day(rs("lasthits"))=day(now()) and month(rs("lasthits"))=month(now()) and year(rs("lasthits"))=year(now()) then
									%><font color=red><%=rs("dayhits")%></font><%
								else
									%><font color=red>0</font><%
								end if
								%>&nbsp;&nbsp;本周：<%
								dim p_year,p_month,p_day,period_time
								p_year=CInt(year(Now()))-CInt(year(rs("lasthits")))
								p_month=CInt(month(Now()))-CInt(month(rs("lasthits")))
								p_day=CInt(day(Now()))-CInt(day(rs("lasthits")))
								period_time=((p_year*12+p_month)*30+p_day)
								if cint(period_time)>=cint(7) then
									%>0<%
								else
									%><%=rs("weekhits")%><%
								end if
							%>&nbsp;&nbsp;总计：<%=rs("hits")%></td>
							<td colspan=2 class=tablebody1 align=center><%=daydownload(rs("dayhits"),rs("daydown"))%></td>
						</tr>
						<%rs.movenext
					loop%>
					<form method=Post action=?classid=<%=request("classid")%>&Nclassid=<%=request("Nclassid")%>&order=<%=request("order")%>&updown=<%=request("updown")%>>
					<tr>
						<td colspan=2 height=25 class=tablebody2>&nbsp;<font color=<%=forum_body(8)%>><b><%=Nclassname%></b></font>&nbsp;&nbsp;页次：<b><%=currentPage%></b> / <b><%=n%></b> 页 每页 <b><%=maxperpage%></b> 软件数 <b><%=totalPut%></b></td>
						<td colspan=3 class=tablebody2 align=right>分页：<%call disppagenum(CurrentPage,n,"?page=","&classid="&request("classid")&"&Nclassid="&request("Nclassid")&"&order="&request("order")&"&updown="&request("updown"))%>&nbsp;转到：<input type=text name=Page size=3 maxlength=10  value=<%=currentpage%>><input type=submit value=Go name=submit>&nbsp;</td>
					</tr>
					</form>
				<%end if%>
			</table>
			<!--#include file=z_down_query.asp-->      
		</td>                           
	</tr>                           
</table>    
<%call down_payuser()
call down_manage()
conndown.close
set conndown=nothing
call activeonline()
call footer()%>