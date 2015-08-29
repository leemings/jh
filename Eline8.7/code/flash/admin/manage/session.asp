<%
if session("UserName")="" then
response.write"<script>alert('登陆超时，请重新登陆！');location='/admin/'</script>"
response.end
end if

if session("level")<>"动画" and session("level")<>"全部" then
response.write"<script>alert('你没有权限进入该管理项目');history.back();</script>"
response.end
end if
%>