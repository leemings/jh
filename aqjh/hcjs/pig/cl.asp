<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_grade<10 or aqjh_name<>application("aqjh_user") then
 Response.Write "<script Language=JavaScript>alert('��ʾ����û��Ȩ�����ý�Ʒ');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql1="delete * from �������ư�"
sql2="delete * from ����ͶƱ��"
connt.execute(sql1)
connt.execute(sql2)
 Response.Write "<script Language=JavaScript>alert('��ʾ���Ѿ������������������Ա���һ�δ����Ŀ�ʼ��');window.location.href ='index.asp';</script>"
 Response.End
%>