<!--#include file="jhb.asp"-->
<%
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_name="" then
%>
<script language="vbscript">
  alert "不能进入，你还没有登录"
  location.href = "../index.asp"
</script>
<%
else
username=sjjh_name
if username<>"root" then%>
<script language="vbscript">
  alert "你没资格操作"
  location.href = "index.asp"
</script>
<%
else
sid=Request.Form ("sid")
sql="update 股票 set 当前价格=当前价格*1.1 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if

