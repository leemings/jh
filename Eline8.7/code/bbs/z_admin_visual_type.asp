<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<%response.buffer=true%>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
	    <th width="10%" height=25">�����ID</th>
	    <th width="*">��������ƣ�����鿴�ӷ��ࣩ</th>
	    <th width="10%">�ӷ������</th>
	    <th width="10%">��Ʒ����</th>
	    <th width="10%">����</th>
	    <th width="30%">����</th>
	  </tr>
		<%if rs.eof and rs.bof then%>
		  <tr> 
		    <td height=25 colspan="6" class=forumrow align=center>û���κδ���࣡</td>
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
			    		%>��ʾ<%
			    	else
			    		%>����<%
			    	end if
			    %></td>
			  	<td class=forumrow><a href="?action=SubTypeAdd&typeid=<%=rs("seqno")%>&typename=<%=rs("typename")%>">����</a> <a href="?action=TypeEdit&typeid=<%=rs("seqno")%>">�༭</a> <a href="?action=TypeDel&typeid=<%=rs("seqno")%>">ɾ��</a> <a href="?action=TypeOrdersUp&Orders=<%=rs("orders")%>">����</a> <a href="?action=TypeOrdersDown&Orders=<%=rs("orders")%>">����</a></td>
			  </tr>
			  <%rs.movenext
			loop
		end if%>
	  <tr> 
	    <td colspan=6 class=forumrow align=right height=25>[<a href="?action=TypeAdd"><b>�����µĴ����</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</td>
	  </tr>
	</table>
	<p align="center"><font color="#FF0000">ע�⣺ɾ��һ������࣬��ߵ�ȫ���ӷ���Ҳ��һ��ɾ������������Ʒ�Ĵ���಻��ɾ����</font></p>
	<%rs.close
	set rs=nothing
end sub

sub TypeAdd
	if not isnull(request("add")) and request("add")<>"" then
		if request("typename")="" then body=body+"<br>"+"��������Ʋ���Ϊ�ա�":exit sub
		dim typeid,orders
		set rs=conn.execute ("select * from Visual_Type where TypeName='"&request("typename")&"'")
		if not (rs.bof and rs.eof) then body=body + "<br>��������Ʋ����ظ������������룡":exit sub
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
	    <th height=25 colspan="2"><a href="?action=TypeInfo"><b>��������</b></a> | �����µĴ����</th>
	  </tr>
		<tr> 
	    <td width="50%" height=25 class=forumrow align=right>���������&nbsp;</td>
	    <td width="50%" class=forumrow align=left>&nbsp;<input name="typename" type="text" size=40></td>
		</tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes" checked> �Ƿ���ʾ�˷���</td>
		</tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="�� ��"></td>
		</tr>
	</table>
	</form>
<%end sub

sub TypeEdit()
	if not isnull(request("edit")) and request("edit")<>"" then
		if request("typename")="" then body=body + "<br>"+"��������Ʋ���Ϊ�գ����������룡":exit sub
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
      <th height=25 colspan="2"><a href="?action=TypeInfo"><b>��������</b></a> | <a href="?action=TypeAdd"><b>�����µĴ����</b></a></th>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>�����ID&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<%=rs("seqno")%></td>
    </tr>
    <tr> 
      <td height=25 width="50%" class=forumrow align=right>���������&nbsp;</td>
      <td width="50%" class=forumrow align=left>&nbsp;<input name="typename" type="text" value=<%=rs("typename")%> size=40></td>
    </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes"<%if rs("visible") then%> checked<%end if%>> �Ƿ���ʾ�˷���</td>
		</tr>
    <tr> 
      <td height=25 colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="�� ��"></td>
    </tr>
	</table>
  </form>
	<%rs.close
	set rs=nothing
end sub

sub TypeDel()
	if conn.execute("select count(*) from visual_items where type mod 100=" & request("typeid"))(0)>0 then  body=body + "<br>"+"��������Ʒ�Ĵ���಻��ɾ����":exit sub
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
	    <th width="10%" height=25>�ӷ���ID</th>
	    <th width="*">�ӷ�������</th>
	    <th width="10%">��Ʒ����</th>
	    <th width="10%">����</th>
	    <th width="30%">����</th>
	  </tr>
	  <%if not rs.eof then
			do while not rs.eof%>
			  <tr align=center> 
			    <td height=25 class=forumrow><font color=<%=forum_body(8)%>><%=rs("seqno")%></font></td>
			    <td class=forumrow><a href="?action=SubTypeEdit&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><%=rs("subtypename")%></a></td>
			    <td class=forumrow><%=conn.execute("select count(*) from visual_items where type=" & (rs("seqno")*100+request("typeid")))(0)%></td>
			    <td class=forumrow><%
			    	if rs("visible") then
			    		%>��ʾ<%
			    	else
			    		%>����<%
			    	end if
			    %></td>
	    		<td class=forumrow><a href="?action=SubTypeEdit&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">�༭</a> <a href="?action=SubTypeDel&subtypeid=<%=rs("seqno")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">ɾ��</a> <a href="?action=OrdersUp&suborders=<%=rs("orders")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">����</a> <a href="?action=OrdersDown&suborders=<%=rs("orders")%>&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>">����</a></td>
	  		</tr>
	  		<%rs.movenext
			loop
		else
			response.write "<tr><td height=25 colspan='4' class=forumrow align=center>Ŀǰ���������л�û���ӷ��࣡</td></tr>"
		end if
		rs.Close
		set rs=nothing%>
	  <tr> 
	    <td height="25" colspan=2 class=forumrow align=left nowrap>&nbsp;&nbsp;<font color=<%=forum_body(8)%>><%=request("typename")%>->�ӷ����б�</b</td>
	    <td colspan=3 class=forumrow align=right>[<a href="?action=TypeInfo"><b>������б�</b></a>] | [<a href="?action=SubTypeAdd&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>�����µ��ӷ���</b></a>]&nbsp;&nbsp;&nbsp;&nbsp;</th>
	  </tr>
	</table>
<%end sub

sub SubTypeAdd()
	if not isnull(request("add")) and request("add")<>"" then
		if request("name")="" then body=body+"<br>"+"�ӷ������Ʋ���Ϊ�ա�":exit sub
		dim subtypeid,orders
		set rs=conn.execute ("select * from Visual_SubType where SubTypeName='"&request("name")&"' and TypeSeq="&request("typeid"))
		if not (rs.bof and rs.eof) then body=body + "<br>�ӷ������Ʋ����ظ������������룡":exit sub
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
			<th colspan=2 height=25><%=request("typename")%>->�����µ��ӷ��� | <a href="?action=SubTypeInfo&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>�ӷ����б�</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>�ӷ�������&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40></td>
	  </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes" checked> �Ƿ���ʾ�˷���</td>
		</tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="�� ��"></td>
	  </tr>
	</table>
	</form>
<%end sub

sub SubTypeEdit()
	if not isnull(request("edit")) and request("edit")<>"" then
		if request("name")="" then body=body + "<br>"+"�ӷ������Ʋ���Ϊ�գ����������룡":exit sub
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
			<th colspan=2 height=25><%=request("typename")%>->�༭�ӷ��� | <a href="?action=SubTypeInfo&typeid=<%=request("typeid")%>&typename=<%=request("typename")%>"><b>�ӷ����б�</b></a></th>
		</tr>
	  <tr> 
	    <td width="40%" height=25 class=forumrow align=right>�ӷ�������&nbsp;</td>
	    <td width="60%" class=forumrow align=left>&nbsp;<input type="text" name="name" size=40 value=<%=rs("subtypename")%>></td>
	  </tr>
		<tr> 
			<td height=25 colspan="2" class=forumrow align=center><input type="checkbox" name="isVisiable" value="yes"<%if rs("visible") then%> checked<%end if%>> �Ƿ���ʾ�˷���</td>
		</tr>
	  <tr> 
	    <td height="25" colspan="2" class=forumrow align=center><input type="submit" name="Submit" value="�� ��"></td>
	  </tr>
	</table>
	</form>
	<%rs.close
	set rs=nothing
end sub

sub SubTypeDel
	if conn.execute("select count(*) from visual_items where type=" & (request("subtypeid")*100+request("typeid")))(0)>0 then  body=body + "<br>"+"��������Ʒ���ӷ��಻��ɾ����":exit sub
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
