<%
if session("UserName")="" then
response.write"<script>alert('��½��ʱ�������µ�½��');location='/admin/'</script>"
response.end
end if

if session("level")<>"����" and session("level")<>"ȫ��" then
response.write"<script>alert('��û��Ȩ�޽���ù�����Ŀ');history.back();</script>"
response.end
end if
%>