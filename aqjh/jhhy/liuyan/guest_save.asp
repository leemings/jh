<!--#include file="conn.asp"-->
<%
user=session("aqjh_name")
if user="" then
response.write"<script>alert('�Բ���ֻ�н�����Ҳſ��������﷢����Ϣ��վ��������ȥ��½����֮��������'); history.back();</script>"
response.end()
end if
Dim user,title,book
title=request("title")
book=request("book")
if title="" or book="" then
response.write"<script>alert('�ܱ�Ǹ������д�ñ���������ٷ�����');history.go(-1);</script>"
response.end()
end if
if len(book)>1000 then
response.write"<script>alert('�Բ��𣬷������ݲ��ܴ��� 1000 ���ַ�');history.go(-1);</script>"
response.end()
end if

set rs=server.createobject("adodb.recordset")
sql="select * from book"
rs.open sql,conn,1,3
rs.addnew
rs("user")=user
rs("title")=title
rs("book")=book
rs.update
rs.close
set rs=nothing
response.write"<script>alert('��ϲ�����ɹ���վ��������һ�����ԣ�ֻ�����Լ����ܿ������������й���Ϣ����ע�����鿴վ�������Ļظ���'); location.replace('chakan.asp');</script>"
response.end()
%>