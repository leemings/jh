<%
for each element in Request.Form
	if Request.Form(element) = "" then 
		Response.write "<META http-equiv=Content-Type content=text/html;charset=gb2312><script>alert('������Ŀ������д������Ϊ��ֵ�������ύ�" & element &"');history.back();</script>"
		Response.End
	end if
Next
%>
