<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftMargin=0 topmargin="2">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	call dvbbs_error()
end if%>
<SCRIPT LANGUAGE=javascript>
function Juge(myform) {
	if (myform.Nclass.value == "") {
		alert("小类名称不能为空！");
		myform.Nclass.focus();
		return (false);
	}
	if (myform.ClassID.value == "") {
		alert("所属大类没有选择！");
		myform.ClassID.focus();
		return (false);
	}
}

function CheckAll(form)  {
	for (var i=0;i<form.elements.length;i++) {
		var e = form.elements[i];
		if (e.name == 'selNclassid') e.checked = !e.checked; 
	}
}
</script>
<%if master then
	select case request("action")
	case "add" 
		call SaveAdd()
	case "modify"  
		call SaveModify()
	case "del"
		call delSubCate()
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
	if Trim(Request.Form("Nclass"))="" or Trim(Request.Form("ClassID"))="" then
		msginfo="添加小类时出现错误，小类名称没有填写或者所属大类没有选择！" 
	else
		set rs=server.createobject("adodb.recordset")   '检查版面是否重名
		rs.open "select * from [DNclass] where Nclass='"&trim(request("Nclass"))&"'",conndown,1,1	
		if not(rs.bof and rs.eof) then
			msginfo="添加小类时出现错误，已经有相同名称的小类已存在！"
		else
			sql="insert into [DNclass](Nclass,ClassID) values('" & trim(request("Nclass")) & "'," & cstr(request("ClassID")) & ")"
			conndown.execute sql  
			msginfo="添加成功，<a href=""?"" class=""ArticleList"">点此继续添加分类</a>！" 
		end if
		rs.close
		set rs=nothing
	end if	
	call Sysmsg("添加小类",msginfo) 
end sub 
  
sub SaveModify()
	dim msginfo
	if Trim(Request.Form("Nclass"))="" or Trim(Request.Form("ClassID"))="" then
		msginfo="修改小类时出现错误，小类名称没有填写或者所属大类没有选择！" 
	else
		set rs=server.createobject("adodb.recordset")   '检查版面是否重名
		rs.open "select * from [DNclass] where Nclass='"&trim(request("Nclass"))&"'",conndown,1,1	
		if not(rs.bof and rs.eof) then
			msginfo="修改小类时出现错误，已经有相同名称的小类已存在！"
		else
			sql="update [DNclass] set Nclass='" &trim(request("Nclass")) & "',ClassID="&request("ClassID")&" where Nclassid=" & request("Nclassid")
			conndown.execute sql
			sql="update [download] set ClassID=" & request("ClassID") & " where Nclassid=" &request("Nclassid")
			conndown.execute sql
			msginfo="修改成功，<a href=""?"" class=""ArticleList"">点此返回小类管理页面</a>！" 
		end if
		rs.close
		set rs=nothing
	end if 		 
	call Sysmsg("修改小类",msginfo) 
end sub 
  
sub delSubCate()
	dim msginfo
	conndown.execute("delete from [DNclass] where Nclassid in ("&Request.Form("selNclassid")&")")
	'删除该分类下的全部软件		
	conndown.execute("delete from [download] where Nclassid in ("&Request.Form("selNclassid")&")") 
	msginfo="删除操作成功，该分类相关信息已全部删除，<a href=""?"">点此返回小类管理页面</a>！" 
	call Sysmsg("删除小类",msginfo) 
end sub

sub myform(isEdit) %> 
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2">管理说明</th>
		</tr>
		<tr>
			<td colspan="2" align="center" class="forumRow"><p align="left"><a href="z_admin_down_class.asp"><font color="#FF0000">管理/添加大分类</font></a>&nbsp;&nbsp;&nbsp;<a href="?"><font color="#FF0000">管理/添加小分类</font></a></td>
		</tr>
	</table>
	<br>
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr>          
			<th><%
				dim boname,SubRs
				set Rs=server.createobject("adodb.recordset")
				set SubRs=server.createobject("adodb.recordset")
				if isedit then
					boname="编辑"
					rs.open "select * from [DNclass] where Nclassid=" & request("Nclassid"),conndown,1,1
					response.write "编辑小类"
				else
					boname="添加"
					response.write "添加小类"
				end if 
			%></th>
		</tr>
		<form name="myform" method="post" action="?" onSubmit="return Juge(this)" >
		<tr>             
			<td align="center" class="forumRow" > <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> <%
				If isedit then
					%><input type="Hidden" name="Nclassid" value='<%=request("Nclassid")%>'><%
				End If
				%><%=boname%>小类名称：<input type="text" name="Nclass" maxlength=40 size="12" value="<%
				if isedit then
					response.write rs("Nclass") 
				end if 
				%>">&nbsp;&nbsp;所属大类：<select name="ClassID"><option value="">请选择...</option><% 
				SubRs.open "select * from [Dclass] order by ClassID",conndown,1,1 
				if not(SubRs.bof and SubRs.eof) then
					do while not SubRs.eof 
						if isedit then 
							%><option value="<%=SubRs("ClassID")%>" <%
							if trim(rs("ClassID"))=trim(SubRs("ClassID")) then
								response.write "selected" 
							end if 
							%>><%=SubRs("Class")%></option><%
						else
							%><option value="<%=SubRs("ClassID")%>"><%=SubRs("Class")%></option><% 
						end if
						SubRs.movenext
					loop
				end if
				SubRs.close
			%></select> <input name="submit" type=submit value="<%=boname%>"></td>
		</tr>
		</form>
		<form name="selform" method="post" action="?" >
		<% if isedit then rs.close
		rs.open "select * from [Dclass] order by ClassID asc",conndown,1,1
		if not(rs.bof and rs.eof) then
			do while not rs.eof %>
				<tr>             
					<td class="forumRowHighlight">&nbsp; <strong><%=rs("Class")%></strong></td>
				</tr>
				<tr> 
					<td class="forumRow" ><%
						SubRs.open "select * from [DNclass] where ClassID="&Rs("ClassID")&" order by Nclassid asc",conndown,1,1
						if SubRs.bof and SubRs.eof then
							response.write "没有添加小类"
						else
							i=1
							do while not SubRs.eof 			 
								%> <input type="checkbox" name="selNclassid" value="<%=SubRs("Nclassid")%>"><a href="?Nclassid=<%=SubRs("Nclassid")%>&action=edit" class="ArticleList"><%=SubRS("Nclass")%></a><%
								if (i mod 7)=0 then response.write "<br>"
								i=i+1
								SubRs.movenext
							loop
						end if
						SubRs.close
					%></td>
				</tr>
				<%rs.movenext
			loop%>
			<tr> 
				<td class="forumRow" align=center>管理操作：全选 <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:CheckAll(this.form)"> <input onClick="{if(confirm('此操作同时将删除该分类下的全部软件！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> <input type="Hidden" name="action" value='del'><font color="#FF0000">此操作同时将删除该分类下的全部软件！</font></td>
			</tr>
		<%end if
		rs.close
		set subrs=nothing
		set rs=nothing%>
		</form>
	</table>
<%end sub 

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
