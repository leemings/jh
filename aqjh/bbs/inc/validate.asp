<%
if pass="0" then
error("<li>本论坛暂时关闭，不再接受访问！")
elseif pass="2" then
if Request.Cookies("username")=empty then error("<li>只有<a href=login.asp>登录</a>后才能浏览论坛")
elseif pass="3" then
if instr("|"&userlist&"|","|"&Request.Cookies("username")&"|")<=0 or Request.Cookies("username")=empty then error("<li>该论坛并未授权您访问！")
elseif pass="4" then
if instr(""&Request.ServerVariables("script_name")&"","newtopic.asp") > 0 and instr("|"&userlist&"|","|"&Request.Cookies("username")&"|")<=0 then error("<li>该论坛并未授权您发帖！")
end if

if password<>empty then
if password<>Request.Cookies("password") then response.redirect "login.asp?menu=password&url=http://"&Request.ServerVariables("server_name")&Request.ServerVariables("script_name")&"?"&Request.ServerVariables("Query_String")&""
end if
%>