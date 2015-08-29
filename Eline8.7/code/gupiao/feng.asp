<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
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
<!--#include file="jhb.asp"-->
<%
for i=1 to 20
sql="update 股票 set 状态='封' where sid="&i
conn.execute sql
next
end if
conn.close
Response.Redirect "index.asp"
%>

