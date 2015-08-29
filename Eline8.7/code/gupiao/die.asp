<%@ LANGUAGE=VBScript%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if
if sjjh_name=""  then
  Response.Redirect "../error.asp?id=440"
elseif sjjh_grade<10 then
	Response.Redirect "../error.asp?id=439"
else%>
<%
sid=Request.Form ("sid")
sql="update 股票 set 当前价格=当前价格*0.9 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if
%>
