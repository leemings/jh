<%
if session("level")<>"全部" then
response.write"<script>alert('你没有权限进入该管理项目');history.back();</script>"
response.end
end if
%>