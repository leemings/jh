<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftMargin=0 topmargin="2">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
end if%>
<SCRIPT LANGUAGE=javascript>
function Juge(myform) {
	if (myform.Nclass.value == "") {
		alert("С�����Ʋ���Ϊ�գ�");
		myform.Nclass.focus();
		return (false);
	}
	if (myform.ClassID.value == "") {
		alert("��������û��ѡ��");
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
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
end if%>
</BODY>

<%sub SaveAdd  
	dim msginfo
	if Trim(Request.Form("Nclass"))="" or Trim(Request.Form("ClassID"))="" then
		msginfo="���С��ʱ���ִ���С������û����д������������û��ѡ��" 
	else
		set rs=server.createobject("adodb.recordset")   '�������Ƿ�����
		rs.open "select * from [DNclass] where Nclass='"&trim(request("Nclass"))&"'",conndown,1,1	
		if not(rs.bof and rs.eof) then
			msginfo="���С��ʱ���ִ����Ѿ�����ͬ���Ƶ�С���Ѵ��ڣ�"
		else
			sql="insert into [DNclass](Nclass,ClassID) values('" & trim(request("Nclass")) & "'," & cstr(request("ClassID")) & ")"
			conndown.execute sql  
			msginfo="��ӳɹ���<a href=""?"" class=""ArticleList"">��˼�����ӷ���</a>��" 
		end if
		rs.close
		set rs=nothing
	end if	
	call Sysmsg("���С��",msginfo) 
end sub 
  
sub SaveModify()
	dim msginfo
	if Trim(Request.Form("Nclass"))="" or Trim(Request.Form("ClassID"))="" then
		msginfo="�޸�С��ʱ���ִ���С������û����д������������û��ѡ��" 
	else
		set rs=server.createobject("adodb.recordset")   '�������Ƿ�����
		rs.open "select * from [DNclass] where Nclass='"&trim(request("Nclass"))&"'",conndown,1,1	
		if not(rs.bof and rs.eof) then
			msginfo="�޸�С��ʱ���ִ����Ѿ�����ͬ���Ƶ�С���Ѵ��ڣ�"
		else
			sql="update [DNclass] set Nclass='" &trim(request("Nclass")) & "',ClassID="&request("ClassID")&" where Nclassid=" & request("Nclassid")
			conndown.execute sql
			sql="update [download] set ClassID=" & request("ClassID") & " where Nclassid=" &request("Nclassid")
			conndown.execute sql
			msginfo="�޸ĳɹ���<a href=""?"" class=""ArticleList"">��˷���С�����ҳ��</a>��" 
		end if
		rs.close
		set rs=nothing
	end if 		 
	call Sysmsg("�޸�С��",msginfo) 
end sub 
  
sub delSubCate()
	dim msginfo
	conndown.execute("delete from [DNclass] where Nclassid in ("&Request.Form("selNclassid")&")")
	'ɾ���÷����µ�ȫ�����		
	conndown.execute("delete from [download] where Nclassid in ("&Request.Form("selNclassid")&")") 
	msginfo="ɾ�������ɹ����÷��������Ϣ��ȫ��ɾ����<a href=""?"">��˷���С�����ҳ��</a>��" 
	call Sysmsg("ɾ��С��",msginfo) 
end sub

sub myform(isEdit) %> 
	<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
		<tr> 
			<th height="23" colspan="2">����˵��</th>
		</tr>
		<tr>
			<td colspan="2" align="center" class="forumRow"><p align="left"><a href="z_admin_down_class.asp"><font color="#FF0000">����/��Ӵ����</font></a>&nbsp;&nbsp;&nbsp;<a href="?"><font color="#FF0000">����/���С����</font></a></td>
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
					boname="�༭"
					rs.open "select * from [DNclass] where Nclassid=" & request("Nclassid"),conndown,1,1
					response.write "�༭С��"
				else
					boname="���"
					response.write "���С��"
				end if 
			%></th>
		</tr>
		<form name="myform" method="post" action="?" onSubmit="return Juge(this)" >
		<tr>             
			<td align="center" class="forumRow" > <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> <%
				If isedit then
					%><input type="Hidden" name="Nclassid" value='<%=request("Nclassid")%>'><%
				End If
				%><%=boname%>С�����ƣ�<input type="text" name="Nclass" maxlength=40 size="12" value="<%
				if isedit then
					response.write rs("Nclass") 
				end if 
				%>">&nbsp;&nbsp;�������ࣺ<select name="ClassID"><option value="">��ѡ��...</option><% 
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
							response.write "û�����С��"
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
				<td class="forumRow" align=center>���������ȫѡ <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:CheckAll(this.form)"> <input onClick="{if(confirm('�˲���ͬʱ��ɾ���÷����µ�ȫ�������\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> <input type="Hidden" name="action" value='del'><font color="#FF0000">�˲���ͬʱ��ɾ���÷����µ�ȫ�������</font></td>
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
			<td height="22" align="center" class="forumRowHighlight"><a href="?" >&lt;&lt; ������һҳ</a></td>
		</tr>
	</table>
	<br>
<%end sub %>
