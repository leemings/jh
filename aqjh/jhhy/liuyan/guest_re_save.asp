<!--#include file="conn.asp"-->
<%
Dim user,re

user=session("aqjh_name")
re=request("re")
if user="" or re="" then
response.write"<script>alert('�Բ��𣬻ظ����ϲ���Ϊ��');history.go(-1);</script>"
response.end()
end if
if len(re)>1000 then
response.write"<script>alert('�Բ��𣬻ظ����ݲ��ܴ��� 1000 ���ַ�');history.go(-1);</script>"
response.end()
end if
set rs=server.createobject("adodb.recordset")
sql="select * from re"
rs.open sql,conn,1,3
rs.addnew
rs("user")=user
rs("re")=re
rs("id")=request("id")
rs.update
rs.close
set rs=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from book where id="&request("id")
rs.open sql,conn,1,3
rs("lastdate")=now()
rs.update
rs.close
set rs=nothing
response.redirect"adminc.asp"
%>