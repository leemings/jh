<%
if session("UserName")="" then
response.write"<script>alert('��½��ʱ�������µ�½��');location='../index.asp'</script>"
response.end
end if
%>