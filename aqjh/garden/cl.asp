<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
sql1="delete * from ��԰���ư�"
sql2="delete * from ��԰ͶƱ��"
connt.execute(sql1)
connt.execute(sql2)
 Response.Write "<script Language=JavaScript>alert('��ʾ���Ѿ����������������Ա���һ�δ����Ŀ�ʼ��');window.location.href ='set.asp';</script>"
 Response.End
%>