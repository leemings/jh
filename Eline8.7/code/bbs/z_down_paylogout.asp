<%
session("payuser")=""
response.redirect Request.ServerVariables("HTTP_REFERER")
%>