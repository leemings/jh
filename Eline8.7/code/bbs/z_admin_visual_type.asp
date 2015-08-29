<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<%response.buffer=true%>
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
else
	dim body
	dim readme,Tlink
	call main()
	set rs=nothing
end if

sub main()
	select case request("action")
	case "TypeInfo"
		call TypeInfo()
	case "TypeAdd"
		call TypeAdd()
	case "TypeEdit"
		call TypeEdit()
	case "TypeDel"
		call TypeDel()
	case "TypeOrdersUp"
		call TypeOrdersUp()
	case "TypeOrdersDown"
		call TypeOrdersDown()
	case "SubTypeInfo"
		call SubTypeInfo()
	case "SubTypeAdd"
		call SubTypeAdd()
	case "SubTypeEdit"
		call SubTypeEdit()
	case "savedit"
		call savedit()
	case "SubTypeDel"
		call SubTypeDel()
	case "OrdersUp"
		call OrdersUp()
	case "OrdersDown"
		call OrdersDown()
	case else
		call TypeInfo()
	end select
	response.write body
end sub

sub TypeInfo()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from [Visual_Type] where visible order by Orders"
	rs.open sql,conn,1,1%>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th width="10%" height=25">大分类ID</th>
	    <th width="*">大分类名称（点击查看子分类）</th>
	    <th width="10%">子分类个数</th>
	    <th width="10%">商品个数</th>
	    <th width="10%">属性</th>
	    <th width="30%">操作</th>
	  </tr>
		<%if rs.eof and rs.bof then%>
		  <tr> 
		    <td height=25 colspan="6" class=forumrow align=center>没有任何大分类！</td>
		  </tr>
		<%else%>
			<%do while not rs.eof%>
			  <tr align=center> 
			    <td height=25 class=forumrow><font color=<%=forum_body(8)%>><%=rs("seqno")%></font></td>
			  	<td class=forumrow><a href="?action=SubTypeInfo&typeid=<%=rs("seqno")%>&typename=<%=rs("typename")%>"><%=rs("typename")%></a></td>
			    <td class=forumrow><%=conn.execute("select count(*) from visual_subtype where typeseq=" & rs("seqno"))(0)%></td>
			    <td class=forumrow><%=conn.execute("select count(*) from visual_items where type mod 100=" & rs("seqno"))(0)%></td>
			    <td class=forumrow><%
			    	if rs("visible") then
			    		%>显示<%
			    	else
			    		%>隐藏<%
			    	end if
			    %></td>
			  	<td class=forumrow><a href="?action=SubTypeAdd&typeid=<%=rs("seqno")%>&typename=<%=rs("typename")%>">增加</a> <a href="?action=TypeEdit&typeid=<%=rs("seqno")%>">编辑</a> <a href="?action=TypeDel&typeid=<%=rs("seqno")%>">删除</a> <a href="?action=TypeOrdersUp&Orders=<%=rs("orders")%>">向上</a> <a href="?action=TypeOrdersDown&Orders=<%=rs("orders")%>">向下</a></td>
			  </tr>
			  <%rs.movenext
			loop
		end if%>
	  <tr> 
	    <td colspan=6 class=forumrow align=right height=25>[<a href="?action=TypeAdd"><b>增加新的大分类</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</td>
	  </tr>
	</table>
	<p align="center"><font color="#FF0000">注意：删除一个大分类，里边的全部子分类也会一起删除，包含有商品的大分类不能删除！</font></p>
	<%rs.close
	set rs=nothing
end sub

sub TypeAdd
	if not isnull(request("add")) and request("add")<>"" then
		if request("typename")="" then body=body+"<br>"+"大分类名称不能为空。":exit sub
		dim typeid,orders
		set rs=conn.execute ("select * from Visual_Type where TypeName='"&request("typename")&"'")
		if not (rs.bof and rs.eof) then body=body + "<br>大分类名称不能重复！请重新输入！":exit sub
		set rs= server.createobject ("adodb.recordset")
		sql = "select top 1 * from Visual_Type order by Seqno desc"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then typeid=1 else typeid=rs("seqno")+1
		rs.close
		sql = "select top 1 * from Visual_Type order by orders desc"
		rs.open sql,conn,1,3
		if rs.bof and rs.eof then orders=1 else orders=rs("orders")+1
		rs.addnew
		rs("seqno")=typeid
		rs("typename")=request("typename")
		rs("orders")=orders
		if request("isVisiable")="yes" then
			rs("visible")=true
		else
			rs("visible")=false
		end if
		rs.update
		rs.close
		set rs=nothing
		response.redirect "?action=TypeInfo"
	end if%>
	<br>
  <form action="?action=TypeAdd&add=1" method = post>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th height=25 colspan="2"><a href="?action=TypeInfo"><b>大分类管理</b></a> | 增加新的大分类</th>
	  </tr>
		<tr> 
	    <td width="50%" height=25 class=forumrow align=right>大分类名称&nbsp;</td>
	    <td width="50%" class=forumrow align=left>&nbsp;<input name="typename" type="text" size=40></td>
		</tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes" checked> 是否显示此分类</td>
		</tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="添 加"></td>
		</tr>
	</table>
	</form>
<%end sub

sub TypeEdit()
	if not isnull(request("edit")) and request("edit")<>"" then
		if request("typename")="" then body=body + "<br>"+"大分类名称不能为空！请重新输入！":exit sub
		if request("isVisiable")="yes" then
			conn.execute "update visual_type set typename='" & request("typename") &"',visible=true where seqno="&request("typeid")
		else
			conn.execute "update visual_type set typename='" & request("typename") &"',visible=false where seqno="&request("typeid")
		end if
		response.redirect "?action=TypeInfo"
	end if
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from visual_type where seqno="&request("typeid")
	rs.open sql,conn,1,1%>
  <br>
  <form action="?action=TypeEdit&edit=1" method=post>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<input type=hidden name=typeid value=<%=Request("typeid")%>>
    <tr> 
      <th height=25 colspan="2"><a href="?action=TypeInfo"><b>大分类管理</b></a> | <a href="?action=TypeAdd"><b>增加新的大分类</b></a></th>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>大分类ID&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<%=rs("seqno")%></td>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>大分类名称&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<input name="typename" type="text" value=<%=rs("typename")%> size=40></td>
    </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes"<%if rs("visible") then%> checked<%end if%>> 是否显示此分类</td>
		</tr>
    <tr> 
      <td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="修 改"></td>
    </tr>
	</table>
  </form>
	<%rs.close
	set rs=nothing
end sub

sub TypeDel()
	if conn.execute("select count(*) from visual_items where type mod 100=" & request("typeid"))(0)>0 then  body=body + "<br>"+"包含有商品的大分类不能删除！":exit sub
	conn.execute "delete * from visual_type where seqno="& request("typeid")
	conn.execute "delete * from visual_subtype where typeseq="& request("typeid")
	response.redirect "?action=TypeInfo"
end sub

sub TypeOrdersUp()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 orders from visual_type where orders<="&request("orders")&" order by orders desc"
	rs.open sql,conn,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("orders")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("orders")
			rs("orders")=oldid
			rs.update
			rs.MovePrevious
			rs("orders")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=TypeInfo"
end sub

sub TypeOrdersDown()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 orders from visual_type where orders>="&request("orders")&" order by orders"
	rs.open sql,conn,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("orders")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("orders")
			rs("orders")=oldid
			rs.update
			rs.MovePrevious
			rs("orders")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=TypeInfo"
end sub

sub SubTypeInfo()
	set rs= server.createobject ("adodb.recordset")
	sql = " select * from visual_subtype where typeseq="&request("typeid")&" order by orders"
	rs.open sql,conn,1,1%> 
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	  <tr> 
	    <th width="10%" height=25>子分类ID</th>
	    <th width="*">子分类名称</th>
	    <th width="10%">商品个数</th>
	    <th width="10%">属性</th>
	    <th width="30%">操作</th>
	  </tr>
	  <%if not rs.eof then
			do while not rs.eof%>
			  <tr align=center> 
			    <td height=25 class=forumrow><font color=<%=forum_body(8)%>><%=rs("seqno")%></font></td>
			    <td class=forumrow><a href="?action=SubTypeEdit&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><%=rs("subtypename")%></a></td>
			    <td class=forumrow><%=conn.execute("select count(*) from visual_items where type=" & (rs("seqno")*100+request("typeid")))(0)%></td>
			    <td class=forumrow><%
			    	if rs("visible") then
			    		%>显示<%
			    	else
			    		%>隐藏<%
			    	end if
			    %></td>
	    		<td class=forumrow><a href="?action=SubTypeEdit&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">编辑</a> <a href="?action=SubTypeDel&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">删除</a> <a href="?action=OrdersUp&suborders=<%=rs("orders")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">向上</a> <a href="?action=OrdersDown&suborders=<%=rs("orders")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">向下</a></td>
	  		</tr>
	  		<%rs.movenext
			loop
		else
			response.write "<tr><td height=25 colspan='4' class=forumrow align=center>目前这个大分类中还没有子分类！</td></tr>"
		end if
		rs.Close
		set rs=nothing%>
	  <tr> 
	    <td height="25" colspan=2 class=forumrow align=left nowrap>&nbsp;&nbsp;<font color=<%=forum_body(8)%>><%=request("typename")%>->子分类列表</b</td>
	    <td colspan=3 class=forumrow align=right>[<a href="?action=TypeInfo"><b>大分类列表</b></a>] | [<a href="?action=SubTypeAdd&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>增加新的子分类</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</th>
	  </tr>
	</table>
<%end sub

sub SubTypeAdd()
	if not isnull(request("add")) and request("add")<>"" then
		if request("name")="" then body=body+"<br>"+"子分类名称不能为空。":exit sub
		dim subtypeid,orders
		set rs=conn.execute ("select * from Visual_SubType where SubTypeName='"&request("name")&"' and TypeSeq="&request("typeid"))
		if not (rs.bof and rs.eof) then body=body + "<br>子分类名称不能重复！请重新输入！":exit sub
		set rs= server.createobject ("adodb.recordset")
		sql = "select top 1 * from Visual_SubType where TypeSeq="&Request("typeid")&" order by Seqno desc"
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then subtypeid=1 else subtypeid=rs("seqno")+1
		rs.close
		sql = "select top 1 * from Visual_SubType where TypeSeq="&request("typeid")&" order by orders desc"
		rs.open sql,conn,1,3
		if rs.bof and rs.eof then orders=1 else orders=rs("orders")+1
		rs.addnew
		rs("seqno")=subtypeid
		rs("typeseq")=request("typeid")
		rs("subtypename")=request("name")
		rs("orders")=orders
		if request("isVisiable")="yes" then
			rs("visible")=true
		else
			rs("visible")=false
		end if
		rs.update
		rs.close
		set rs=nothing
		response.redirect "?action=SubTypeInfo&typeid="&request("typeid")&"&typename="&request("typename")
	end if%>
	<br>
	<form action="?action=SubTypeAdd&add=1" method = post>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<input type="hidden" name="typeid" value=<%=request("typeid")%>>
		<input type="hidden" name="typename" value=<%=request("typeName")%>>
		<tr> 
			<th colspan=2 height=25><%=request("typename")%>->增加新的子分类 | <a href="?action=SubTypeInfo&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>子分类列表</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>子分类名称&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40></td>
	  </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes" checked> 是否显示此分类</td>
		</tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="添 加"></td>
	  </tr>
	</table>
	</form>
<%end sub

sub SubTypeEdit()
	if not isnull(request("edit")) and request("edit")<>"" then
		if request("name")="" then body=body + "<br>"+"子分类名称不能为空！请重新输入！":exit sub
		if request("isVisiable")="yes" then
			conn.execute "update visual_subtype set subtypename='" & request("name") &"',visible=true where seqno="&request("subtypeid")&" and typeseq="&request("typeid")
		else
			conn.execute "update visual_subtype set subtypename='" & request("name") &"',visible=false where seqno="&request("subtypeid")&" and typeseq="&request("typeid")
		end if
		response.redirect "?action=SubTypeInfo&typeid="&request("typeid")&"&typename="&request("typename")
	end if
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from visual_subtype where seqno="&request("subtypeid")&" and typeseq="&request("typeid")
	rs.open sql,conn,1,1%>
	<br>
	<form action="?action=SubTypeEdit&edit=1" method=post>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
		<input type="hidden" name="typeid" value=<%=request("typeid")%>>
		<input type="hidden" name="typename" value=<%=request("typeName")%>>
		<input type="hidden" name="subtypeid" value=<%=Request("subtypeid")%>>
		<tr> 
			<th colspan=2 height=25><%=request("typename")%>->编辑子分类 | <a href="?action=SubTypeInfo&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>子分类列表</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>子分类名称&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40 value=<%=rs("subtypename")%>></td>
	  </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes"<%if rs("visible") then%> checked<%end if%>> 是否显示此分类</td>
		</tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="修 改"></td>
	  </tr>
	</table>
	</form>
	<%rs.close
	set rs=nothing
end sub

sub SubTypeDel
	if conn.execute("select count(*) from visual_items where type=" & (request("subtypeid")*100+request("typeid")))(0)>0 then  body=body + "<br>"+"包含有商品的子分类不能删除！":exit sub
	conn.Execute("delete from visual_subtype where seqno="&request("subtypeid")&" and typeseq="&request("typeid"))
	response.redirect "?action=SubTypeInfo&typeid="&request("typeid")&"&typename="&request("typename")
end sub

sub OrdersUp()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 orders from visual_subtype where orders<="&request("suborders")&" and typeseq="&request("typeid")&" order by orders desc"
	rs.open sql,conn,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("orders")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("orders")
			rs("orders")=oldid
			rs.update
			rs.MovePrevious
			rs("orders")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=SubTypeInfo&typeid="&request("typeid")&"&typename="&request("typename")
end sub

sub OrdersDown()
	dim oldid,newid
	set rs= server.createobject ("adodb.recordset")
	sql="select top 2 orders from visual_subtype where orders>="&request("suborders")&" and typeseq="&request("typeid")&" order by orders"
	rs.open sql,conn,1,3
	if not (rs.eof or rs.bof) then
		oldid=rs("orders")
		rs.movenext
		if not (rs.eof or rs.bof) then
			newid=rs("orders")
			rs("orders")=oldid
			rs.update
			rs.MovePrevious
			rs("orders")=newid
			rs.update
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "?action=SubTypeInfo&typeid="&request("typeid")&"&typename="&request("typename")
end sub%>
