<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
<SCRIPT LANGUAGE=javascript>
function Juge(myform) {
	if (myform.Class.value == "") {
		alert("�������Ʋ���Ϊ�գ�");
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
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
end if%>
</BODY>  

<%sub SaveAdd
	dim msginfo  
	if trim(request("Class"))="" then
		msginfo="���ʱ���ִ��󣬴������Ʋ���Ϊ�գ�" 
	else
		set rs=server.createobject("adodb.recordset")   '�������Ƿ�����
		rs.open "select * from Dclass where Class='" & trim(request("Class")) & "'",conndown,1,1	     
		if not(rs.bof and rs.eof) then
			msginfo="���ʱ���ִ����Ѿ�����ͬ���ƵĴ����Ѵ��ڣ�"
		else
			sql="insert into [Dclass](Class) values('" & trim(request("Class")) & "')"
			conndown.execute sql  
			msginfo="��ӳɹ���<a href=""?"">��˼�����ӷ���</a>��" 
		end if
		rs.close
		set rs=nothing
	end if	
	call Sysmsg("��Ӵ���",msginfo) 
end sub 
  
sub SaveModify()
	dim msginfo  
	if trim(request("Class"))="" then
		msginfo="�޸Ĵ���ʱ���ִ��󣬴������Ʋ���Ϊ�գ�" 
	else
		set rs=server.createobject("adodb.recordset")   '�������Ƿ�����
		rs.open "select * from [Dclass] where Class='" & trim(request("Class")) & "'",conndown,1,1	     
		if not(rs.bof and rs.eof) then
			msginfo="�޸Ĵ���ʱ���ִ����Ѿ�����ͬ���ƵĴ����Ѵ��ڣ�"
		else
			sql="update [Dclass] set Class='" & trim(request("Class")) & "' where ClassID=" & cstr(request("classID"))
			conndown.execute sql
			msginfo="�޸ĳɹ���<a href=""?"">��˷��ش������ҳ��</a>��" 
		end if
		rs.close
		set rs=nothing
	end if 		 
	call Sysmsg("�޸Ĵ���",msginfo) 
end sub 
  
sub delCate()
	dim msginfo  
	conndown.execute("delete from [Dclass] where ClassID in ("&Request.Form("selCateID")&")")
	'ɾ���÷����µ�ȫ���ӷ��༰���
	conndown.execute("delete from [DNclass] where classID in ("&Request.Form("selCateID")&")")		
	conndown.execute("delete from [download] where classID in ("&Request.Form("selCateID")&")") 
	msginfo="ɾ�������ɹ����÷��������Ϣ��ȫ��ɾ����<a href=""?"">��˷��ش������ҳ��</a>��" 
	call Sysmsg("ɾ������",msginfo) 
end sub

sub myform(isEdit)
	dim boname%>
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2">����˵��</th>
		</tr>
		<tr>
			<td colspan="2" align="center" class="forumRow"><p align="left"><a href="?"><font color="#FF0000">����/��Ӵ����</font></a>&nbsp;&nbsp;&nbsp;<a href="z_admin_down_subclass.asp"><font color="#FF0000">����/���С����</font></a></td>
		</tr>
	</table>
	<br>
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2"><%
				set rs=server.createobject("adodb.recordset")
				if isedit then
					boname="�޸�"
					rs.open "select * from [Dclass] where ClassID=" & request("ClassID"),conndown,1,1
					response.write "�༭����"
				else
					response.write "��Ӵ���"
					boname="���"
				end if 
			%></th>
		</tr>
		<form name="myform" method="post" action="?" onSubmit="return Juge(this)" >
		<tr> 
			<td colspan="2" align="center" class="forumRow"><input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'><%
				If isedit then
					%> <input type="Hidden" name="ClassID" value='<%=request("ClassID")%>'><%
				End If
				%><%=boname%>�������ƣ�<input type="text" name="Class" maxlength=40 size="20" value='<%
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
			<td width="85%" align="center" class="forumRowHighlight">��������</td>
		</tr>
		<%rs.open "select * from [Dclass] order by ClassID asc",conndown,1,1
		if rs.bof and rs.eof then
			response.write "Ŀǰû���κΰ���"
		else 
			do while not rs.eof %>
				<tr> 
					<td width="15%" align="center" class="forumRow" ><%=rs("classID")%></td>
					<td width="85%" class="forumRow" > <input name="selCateID" type="checkbox" value="<%=rs("classID")%>"> <a href="?classID=<%=rs("classID")%>&action=edit" class="ArticleList"><%=rs("Class")%></a></td>
				</tr>
				<% rs.movenext
			loop%>
			<tr> 
				<td colspan=2 align=center class="forumRow">���������ȫѡ <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:CheckAll(this.form)"> <input onClick="{if(confirm('�˲���ͬʱ��ɾ���÷����µ�ȫ���ӷ��༰�����\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> <input type="Hidden" name="action" value='del'> <font color="#FF0000">�˲���ͬʱ��ɾ���÷����µ�ȫ���ӷ��༰���</font></td>
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
			<td height="22" align="center" class="forumRowHighlight"><a href="?" >&lt;&lt; ������һҳ</a></td>
		</tr>
	</table>
	<br> 
<%end sub %> 
