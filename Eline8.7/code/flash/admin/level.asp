<%
if session("level")<>"ȫ��" then
response.write"<script>alert('��û��Ȩ�޽���ù�����Ŀ');history.back();</script>"
response.end
end if
%>