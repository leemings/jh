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
name=request("name")
j=request("j")
if name="" or j="" then
 Response.Write "<script Language=JavaScript>alert('��ʾ������!');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * from ��԰����",connt,3,3
jz=rs("ͶƱ����")
if cdate(jz)>now() then
 Response.Write "<script Language=JavaScript>alert('��ʾ��ͶƱû���������ܰ佱��');window.close();</script>"
 Response.End
end if
rs.close
sql="Select * from ��԰���ư� where ������='"&name&"'"
rs.open sql,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('��ʾ�����˲�û�в�������');window.close();</script>"
 Response.End
else
   if rs("����")>0 and j<>0 then
       Response.Write "<script Language=JavaScript>alert('��ʾ�����ܵڶ��ΰ佱��');window.close();</script>"
       Response.End
   else
   select case j
   case 0
    sql2="update ��԰���ư� set ����=0 where ������='"&name&"'"
    connt.execute(sql2)
    Response.Write "<script Language=JavaScript>alert('��ʾ��"&name&"�����Ѿ������');window.close();</script>"
    Response.End
   case 1,2,3
    sql2="update ��԰���ư� set ����="&j&",��ʱ��='"&now()&"' where ������='"&name&"'"
    connt.execute(sql2)
    Response.Write "<script Language=JavaScript>alert('�佱��"&name&"�Ѿ�����Ϊ"&j&"�Ƚ���');window.close();</script>"
    Response.End
   case else
       Response.Write "<script Language=JavaScript>alert('��ʾ������');window.close();</script>"
       Response.End
   end select
   end if
end if
%>