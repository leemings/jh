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
  alert "���ܽ��룬�㻹û�е�¼"
  location.href = "../index.asp"
</script>
<%
else
username=sjjh_name
if username<>"root" then%>
<script language="vbscript">
  alert "��û�ʸ����"
  location.href = "index.asp"
</script>
<%
else
sid=Request.Form ("sid")
sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*1.1 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if

