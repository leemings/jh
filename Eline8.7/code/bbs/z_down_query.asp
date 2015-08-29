<form method="post" action="z_down_ru_query.asp">软件搜索：
	<select name="action" size="1">
	<option value="title">程序名称</option>
	<option value="content">程序简介</option>
	</select>
	<select name="classid" size="1">
	<%set rs=server.createobject("adodb.recordset")
	sql="select * from Dclass"
	rs.open sql,conndown,1,1
	if rs.bof and rs.eof then
		%><option value="">未指定条件</option><%
	else
		%><option value="" selected>未指定条件</option><%
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
		%><option value="" selected>未指定条件</option><%
	else
		%><option value="" selected>未指定条件</option><%
		do while not rs.eof
			%><option value="<%=rs("Nclassid")%>"><%=rs("Nclass")%></option><%
			rs.movenext
		loop
	end if%>
	</select>
	<input type="text" name="keyword" class=smallinput size=10 value="关键字" maxlength="50" onfocus="this.select();">
	<input type="submit" name="Submit" value="搜索" class="buttonface">
</form>