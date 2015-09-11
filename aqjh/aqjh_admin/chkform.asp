<%
for each element in Request.Form
	if Request.Form(element) = "" then 
		Response.write "<META http-equiv=Content-Type content=text/html;charset=gb2312><script>alert('所有项目必须填写，不可为空值，错误提交项：" & element &"');history.back();</script>"
		Response.End
	end if
Next
%>
