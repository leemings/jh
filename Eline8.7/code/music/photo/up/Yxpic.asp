<!--#include file="conn.asp"-->

<%

m=0
set rs=conn.execute("SELECT top 1 ID FROM images order by ID desc") 
if not Rs.eof then
	do while not rs.eof
%>
<html>
<head>
<title>文件上传</title>
</head>
<script>parent.form2.pic.value+='showimg.asp?id=<%=rs("ID")%>'</script>
<body leftmargin=0 topmargin=0>showimg.asp?id=<%=rs("ID")%>
<%
		if m>=1 then exit do
	rs.movenext   
	loop
end if
rs.close
set rs=nothing
set conn=nothing
%>
