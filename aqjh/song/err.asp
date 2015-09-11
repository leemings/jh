<%
	response.write "<script>alert('提示:你输入的祝福语中有系统禁止脏话!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
%>