<form method="post" action="z_down_ru_query.asp">���������
	<select name="action" size="1">
	<option value="title">��������</option>
	<option value="content">������</option>
	</select>
	<select name="classid" size="1">
	<%set rs=server.createobject("adodb.recordset")
	sql="select * from Dclass"
	rs.open sql,conndown,1,1
	if rs.bof and rs.eof then
		%><option value="">δָ������</option><%
	else
		%><option value="" selected>δָ������</option><%
		do while not rs.eof
			%><option value="<%=rs("classid")%>"><%=rs("class")%></option><%
			rs.movenext
		loop
	end if
	rs.close%>
	</select>
	<select name="Nclassid" size="1">
	<%sql="select * from DNclass"
	rs.open sql,conndown,1,1
	if rs.bof and rs.eof then
		%><option value="" selected>δָ������</option><%
	else
		%><option value="" selected>δָ������</option><%
		do while not rs.eof
			%><option value="<%=rs("Nclassid")%>"><%=rs("Nclass")%></option><%
			rs.movenext
		loop
	end if%>
	</select>
	<input type="text" name="keyword" class=smallinput size=10 value="�ؼ���" maxlength="50" onfocus="this.select();">
	<input type="submit" name="Submit" value="����" class="buttonface">
</form>