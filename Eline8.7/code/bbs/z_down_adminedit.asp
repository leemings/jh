<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file=z_down_function.asp-->
<%dim MaxPerPage,currentpage,TitleName,Classid,Nclassid
stats="修改删除软件"
call nav()
call head_var(0,0,"下载中心","z_down_default.asp")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您没有进入下载中心管理的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
elseif flagdown>2 or flagdown="" or flagdown<=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有管理软件的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	call dvbbs_error()
else
	MaxPerPage=editnum
	if not isempty(request("page")) then
		currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	set Rs=server.createobject("adodb.recordset")
	if request("Classid")<>"" and request("Nclassid")<>"" then
		Classid=" Classid="&request("Classid")&" and "
		Nclassid=" Nclassid="&request("Nclassid")&" "
		sql="select * from DNclass where Nclassid="&request("Nclassid")
		rs.open sql,conndown,1,1
		TitleName=trim(rs("Nclass"))
		rs.close
	end if
	if request("Classid")<>"" and request("Nclassid")="" then
		Classid="Classid="&request("Classid")&" "
		sql="select Class from Dclass where Classid="&request("Classid")
		rs.open sql,conndown,1,1
		TitleName=trim(rs("Class"))
		rs.close
	end if
	if request("Classid")="" and request("Nclassid")="" then
		Classid=""
		Nclassid=""
		TitleName="全部软件"
	end if%>
	<!-- #include file="z_down_forum_js.asp" -->
	<table cellpadding=3 cellspacing=1 border=0 width=<%=forum_body(12)%> align=center>
		<tr>
			<td width="100%" align="center" valign=top class=tablebody2><%
				rs.open "select * from [Dclass] order by Classid asc",conndown,1,1
				if rs.bof and rs.eof then
					response.write "没有添加大类"
				else
					do while not rs.eof
						Response.Write " <a href=""?Classid="&rs("Classid")&""" onMouseOver='ShowMenu(catemenunew"&rs("Classid")&",120)'>"&rs("Class")&"</a>" & vbcrlf
						rs.movenext
					loop
				end if
				rs.close
			%></td>
		</tr>
	</table>
	<%if flagdown<=2 then
		if not isempty(Request.Form("selID")) then
			dim selID
			selid=replace(request.form("selID"),"'","")
			selid=replace(selid,";","")
			selid=replace(selid,"--","")
			selid=replace(selid,")","")
			founderr=false
			select case Trim(Request.Form("action_type"))
			case "isdel"
				call isdel()
			case "ismove"
				call ismove()
			case "setSoftType"
				call setSoftType()
			case "isCommend"
				call isCommend()
			case "noCommend"
				call noCommend()
			case "ishide"
				call ishide()
			case "nohide"
				call nohide()
			case else
				founderr=true
			end select
			if not founderr then
				call dvbbs_suc()
			end if
		else
			call SoftList()
		end if
	else
		call SoftList()
	end if
	set rs=nothing
end if
conndown.close
set conndown=nothing	
call activeonline()
call footer()

sub SoftList()%>
	<table cellpadding="3" cellspacing="1" border="0" class="tableBorder1" align=center>
		<form name="myform" method="post" action="?">
		<tr> 
			<th height="22">软件名称</th>
			<th>是否隐藏</th>
			<th>推荐</th>
			<th>软件类型</th>
			<th>整理时间</th>
		</tr>
		<%if request("Classid")<>"" and request("Nclassid")<>"" then
			sql="select * from [download] where "&Classid&" "&Nclassid&" "
			sql=sql&" order by id desc"
		end if
		if request("Classid")<>"" and request("Nclassid")="" then
			sql="select * from [download] where "&Classid&" "
			sql=sql&" order by id desc"
		end if
		if request("Classid")="" and request("Nclassid")="" then
			sql="select * from [download] "
			sql=sql&" order by id desc"
		end if
		rs.open sql,conndown,1,1 
		if rs.eof or rs.bof then
			response.write "<tr height=25><td class=tablebody1 align=center colspan=5>没有或未找到任何程序</td></tr>"
		else
			dim totalput,n
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
					<td class="tablebody1"><%
						if flagdown<=2 then
							%><input type="checkbox" name="selID" value="<%=rs("ID")%>"><%
						end if
					%><a href="z_down_editsoft.asp?id=<%=rs("id")%>&classid=<%=rs("classid")%>&Nclassid=<%=rs("Nclassid")%>"><%=rs("showname")%></a></td>
					<td align="center" class="tablebody1"><%
						if rs("stop")=1 then
							response.write "隐藏"
						else 
							response.write "正常" 
						end if 
					%></td>
					<td align="center" class="tablebody1"><%
						if rs("hots")=1 then
							response.write "<font color=red><b>√</b></font>"
						else
							response.write "<b>×</b>"
						end if
					%></td>
					<td align="center" class="tablebody1"><%=downshow(rs("downshow"))%></td>
					<td align="center" class="tablebody1"><%
						if rs("dateandtime")=Date() then
							response.write "<font color=red>"&rs("dateandtime")&"</font>"
						else
							response.write rs("dateandtime")
						end if
					%></td>
				</tr>
				<%rs.movenext
			loop
			rs.close%>
			<tr>
				<td height="25" class=tablebody2 align=right colspan=5>页次：<b><%=currentPage%></b> / <b><%=n%></b> 页 每页 <b><%=maxperpage%></b> 软件数 <b><%=totalPut%></b>&nbsp;&nbsp;&nbsp;&nbsp;分页：<%call disppagenum(currentpage,n,"?page=","&classid="&request("classid")&"&Nclassid="&request("Nclassid"))%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
			</tr>
			<%if flagdown<=2 then
				sql = "select * from DNclass order by Nclassid asc"
				rs.open sql,conndown,1,1%>
				<SCRIPT language = "JavaScript">
					var onecount;
					onecount=0;
					subcat = new Array();
					<%dim count
					count = 0
					do while not rs.eof %>
						subcat[<%=count%>] = new Array("<%= trim(rs("Nclass"))%>","<%=cstr(rs("Classid"))%>","<%=cstr(rs("Nclassid"))%>");
						<%count = count + 1
						rs.movenext
					loop
					rs.close%>
					onecount=<%=count%>;
		
					function changelocation(locationid)	{
						document.myform.selNclassid.length = 0; 
						var locationid=locationid;
						var i;
						for (i=0;i < onecount; i++) {
							if (subcat[i][1] == locationid) { 
				 				document.myform.selNclassid.options[document.myform.selNclassid.length] = new Option(subcat[i][0], subcat[i][2]);
							} 
						}
					} 
		
					function CheckAll(form)  {
						for (var i=0;i<form.elements.length;i++) {
							var e = form.elements[i];
							if (e.name == 'selID') e.checked = !e.checked; 
						}
					}
				</script>
				<tr>
					<td colspan="5" class="tablebody1">&nbsp;&nbsp;全选：<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">&nbsp;<select name="selClassid" onChange="changelocation(document.myform.selClassid.options[document.myform.selClassid.selectedIndex].value)"><%
						sql="select * from Dclass"
						rs.open sql,conndown,1,1
						if rs.eof and rs.bof then
							response.write "<option value=""0"">没有大类</option>"
						else
							do while not rs.eof
								response.write "<option value="""&cstr(rs("Classid"))&""">"&trim(rs("Class"))&"</option>"
								rs.movenext
							loop
						end if
						rs.close
						%></select>&nbsp;<select name="selNclassid"><% 
						sql="select * from DNclass"
						rs.open sql,conndown,1,1
						if rs.eof and rs.bof then
							response.write "<option value=""0"">没有小类</option>"
						else
							do while not rs.eof
								response.write "<option value="""&cstr(rs("Nclassid"))&""">" & trim(rs("Nclass")) & "</option>"
								rs.MoveNext
							Loop
						end if
						rs.close
						%></select>&nbsp;<select name="downshow"><option value="">下载属性</option><option value="1">开放下载</option><option value="2">会员下载</option><%
						if vipshow=2 then
							%><option value="3">VIP下载</option><%
						end if%><option value="4">特约下载</option></select><br>&nbsp;&nbsp;
						<input type="radio" name="action_type" value="isdel">批量删除 
						<input type="radio" name="action_type" value="ismove">批量移动 
						<input type="radio" name="action_type" value="setSoftType">设置下载属性 
						<input type="radio" name="action_type" value="isCommend">批量推荐 
						<input type="radio" name="action_type" value="noCommend">取消推荐
						<input type="radio" name="action_type" value="ishide">批量隐藏
						<input type="radio" name="action_type" value="nohide">取消隐藏
						<input type="submit" name="Submit" value="执行" onclick="{if(confirm('您确定执行此操作吗?')){this.document.myform.submit();return true;}return false;}">
					</td>
				</tr>
			<%end if
		end if%>
		</form>
	</table>
<%end sub 

sub ishide()
	conndown.execute("update [download] set stop=1 where ID in ("&selID&")")
	sucmsg="<li>批量隐藏操作成功！"
end sub

sub nohide()
	conndown.execute("update [download] set stop=0 where ID in ("&selID&")")
	sucmsg="<li>批量取消隐藏操作成功！"
end sub


sub isdel()
	conndown.execute("delete from [download] where ID in ("&selID&")")
	sucmsg="<li>批量删除操作成功！"
end sub

sub ismove()
	if Trim(Request.Form("selClassid"))<>"" and Trim(Request.Form("selNclassid"))<>"" then
		conndown.execute("update [download] set Classid='"&Trim(Request.Form("selClassid"))&"',Nclassid='"&Trim(Request.Form("selNclassid"))&"' where ID in ("&selID&")")
		sucmsg="<li>批量删除操作成功！"
	end if
end sub

Sub setSoftType()
	if Trim(Request.Form("downshow"))<>"" then
		conndown.execute("update [download] set downshow="&Request.Form("downshow")&" where ID in ("&selID&")")
		sucmsg="<li>批量设置软件下载属性操作成功！"
	end if
end Sub

sub isCommend()
	conndown.execute("update [download] set hots=1 where ID in ("&selID&")")
	sucmsg="<li>批量推荐软件操作成功！"
end sub

sub noCommend()
	conndown.execute("update [download] set hots=0 where ID in ("&selID&")")
	sucmsg="<li>批量取消推荐软件操作成功！"
end sub%>
