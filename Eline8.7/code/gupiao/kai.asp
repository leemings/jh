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
alert "������������ٽ����Ʊ��"
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
sql= "select * from ��Ʊ where sid=1"        
set rss=conn.execute(sql)
if rss("״̬")="��" then 
%>
<script language="vbscript">
  alert "Ŀǰ�����Ѿ��ǿ���״̬��"
  location.href = "index.asp"
</script>
<%
elseif rss("����")=date() then
%>
<script language="vbscript">
  alert "��������Ѿ����̲����ٿ���"
  location.href = "index.asp"
</script>
<%
else
for i=1 to 20  
sql= "select * from ��Ʊ where sid="&i        
set rs=conn.execute(sql)   
sql="update ��Ʊ set ���̼۸�='"&rs("��ǰ�۸�")&"',״̬='��',����='"&date()&"'where sid="&i
conn.execute sql
next
end if
end if
conn.close
Response.Redirect "index.asp"

%>

