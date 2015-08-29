<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
<SCRIPT LANGUAGE=javascript>
function Juge(myform) {
	if (myform.Class.value == "") {
		alert("大类名称不能为空！");
		myform.Class.focus();
		return (false);
	}
}

function CheckAll(form)  {
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name == 'selCateID') e.checked = !e.checked; 
	}
}
</script>
</head>
<BODY leftMargin=0>
<%if master then   
	select case request("action")
	case "add"
		call SaveAdd()
	case "modify"
		call SaveModify()
	case "del"
		call delCate()
	case "edit"
		call myform(True)
	case else
		call myform(False) 
	end select
else
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
end if%>
</BODY>  

<%sub SaveAdd
	dim msginfo  
	if trim(request("Class"))="" then
		msginfo="添加时出现错误，大类名称不能为空！" 
	else
		set rs=server.createobject("adodb.recordset")   '检查版面是否重名
		rs.open "select * from Dclass where Class='" & trim(request("Class")) & "'",conndown,1,1	     
		if not(rs.bof and rs.eof) then
			msginfo="添加时出现错误，已经有相同名称的大类已存在！"
		else
			sql="insert into [Dclass](Class) values('" & trim(request("Class")) & "')"
			conndown.execute sql  
			msginfo="添加成功，<a href=""?"">点此继续添加分类</a>！" 
		end if
		rs.close
		set rs=nothing
	end if	
	call Sysmsg("添加大类",msginfo) 
end sub 
  
sub SaveModify()
	dim msginfo  
	if trim(request("Class"))="" then
		msginfo="修改大类时出现错误，大类名称不能为空！" 
	else
		set rs=server.createobject("adodb.recordset")   '检查版面是否重名
		rs.open "select * from [Dclass] where Class='" & trim(request("Class")) & "'",conndown,1,1	     
		if not(rs.bof and rs.eof) then
			msginfo="修改大类时出现错误，已经有相同名称的大类已存在！"
		else
			sql="update [Dclass] set Class='" & trim(request("Class")) & "' where ClassID=" & cstr(request("classID"))
			conndown.execute sql
			msginfo="修改成功，<a href=""?"">点此返回大类管理页面</a>！" 
		end if
		rs.close
		set rs=nothing
	end if 		 
	call Sysmsg("修改大类",msginfo) 
end sub 
  
sub delCate()
	dim msginfo  
	conndown.execute("delete from [Dclass] where ClassID in ("&Request.Form("selCateID")&")")
	'删除该分类下的全部子分类及软件
	conndown.execute("delete from [DNclass] where classID in ("&Request.Form("selCateID")&")")		
	conndown.execute("delete from [download] where classID in ("&Request.Form("selCateID")&")") 
	msginfo="删除操作成功，该分类相关信息已全部删除，<a href=""?"">点此返回大类管理页面</a>！" 
	call Sysmsg("删除大类",msginfo) 
end sub

sub myform(isEdit)
	dim boname%>
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2">管理说明</th>
		</tr>
		<tr>
			<td colspan="2" align="center" class="forumRow"><p align="left"><a href="?"><font color="#FF0000">管理/添加大分类</font></a>&nbsp;&nbsp;&nbsp;<a href="z_admin_down_subclass.asp"><font color="#FF0000">管理/添加小分类</font></a></td>
		</tr>
	</table>
	<br>
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2"><%
				set rs=server.createobject("adodb.recordset")
				if isedit then
					boname="修改"
					rs.open "select * from [Dclass] where ClassID=" & request("ClassID"),conndown,1,1
					response.write "编辑大类"
				else
					response.write "添加大类"
					boname="添加"
				end if 
			%></th>
		</tr>
		<form name="myform" method="post" action="?" onSubmit="return Juge(this)" >
		<tr> 
			<td colspan="2" align="center" class="forumRow"><input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'><%
				If isedit then
					%> <input type="Hidden" name="ClassID" value='<%=request("ClassID")%>'><%
				End If
				%><%=boname%>大类名称：<input type="text" name="Class" maxlength=40 size="20" value='<%
				if isedit then
					response.write rs("Class") 
					rs.close
				end if 
			%>'> <input name="submit" type=submit value="<%=boname%>"></td>
		</tr>
		</form>
		<form name="selform" method="post" action="?" >
		<tr> 
			<td width="15%" align="center" class="forumRowHighlight">ID</td>
			<td width="85%" align="center" class="forumRowHighlight">版面名称</td>
		</tr>
		<%rs.open "select * from [Dclass] order by ClassID asc",conndown,1,1
		if rs.bof and rs.eof then
			response.write "目前没有任何版面"
		else 
			do while not rs.eof %>
				<tr> 
					<td width="15%" align="center" class="forumRow" ><%=rs("classID")%></td>
					<td width="85%" class="forumRow" > <input name="selCateID" type="checkbox" value="<%=rs("classID")%>"> <a href="?classID=<%=rs("classID")%>&action=edit" class="ArticleList"><%=rs("Class")%></a></td>
				</tr>
				<% rs.movenext
			loop%>
			<tr> 
				<td colspan=2 align=center class="forumRow">管理操作：全选 <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:CheckAll(this.form)"> <input onClick="{if(confirm('此操作同时将删除该分类下的全部子分类及软件！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> <input type="Hidden" name="action" value='del'> <font color="#FF0000">此操作同时将删除该分类下的全部子分类及软件</font></td>
			</tr>
		<%end if
		rs.close
		set rs=nothing%>
		</form>
	</table>
<% end sub 

sub Sysmsg(msgtitle,msginfo)%>
	<br>
	<table width="85%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder" >
		<tr> 
			<th><%=msgtitle%></th>
		</tr>
		<tr>           
			<td class="forumRow"><%=msginfo%></td>
		</tr>
		<tr>           
			<td height="22" align="center" class="forumRowHighlight"><a href="?" >&lt;&lt; 返回上一页</a></td>
		</tr>
	</table>
	<br> 
<%end sub %> 
