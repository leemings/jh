<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id FROM �û� WHERE  ����='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
name=trim(request("name"))
if name=aqjh_name then
 Response.Write "<script Language=JavaScript>alert('��ʾ��Ϊ�����ף����ø��Լ�ͶƱ��');window.close();</script>"
 Response.End
end if
rs.close
%>
<!--#include file="data.asp"-->
<%
rs.open "Select * from ��Ȧ����",connt,3,3
ks=rs("ͶƱ��ʼ")
jz=rs("ͶƱ����")
dj=rs("�ȼ�")
rs.close
if aqjh_jhdj<dj then
 Response.Write "<script Language=JavaScript>alert('��ʾ���ȼ�����������ͶƱ��');window.close();</script>"
 Response.End
end if
if cdate(ks)>now() then
 Response.Write "<script Language=JavaScript>alert('��ʾ��ͶƱû��ʼ�����ܰ佱��');window.close();</script>"
 Response.End
end if
if cdate(jz)<now() then
 Response.Write "<script Language=JavaScript>alert('��ʾ��ͶƱû���������ܰ佱��');window.close();</script>"
 Response.End
end if
sql1="select * from ������ư� where ������='"&name&"'"
rs.open sql1,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('��ʾ������δ�������������');window.close();</script>"
 Response.End
else
 rs.close
 sql2="select * from ����ͶƱ�� where ͶƱID="&aqjh_id
 rs.open sql2,connt,1,1
 if rs.recordcount<>0 then
   Response.Write "<script Language=JavaScript>alert('��ʾ�������ظ�ͶƱ��');window.close();</script>"
   Response.End
 else
   sql3="update ������ư� set Ʊ��=Ʊ��+1 where ������='"&name&"'"
   connt.execute(sql3)
   sql4="INSERT INTO ����ͶƱ�� (ͶƱID,����) VALUES ("&aqjh_id&",'"&aqjh_name&"')"
   connt.execute(sql4)
   Response.Write "<script Language=JavaScript>alert('��ʾ��лл��Ͷ�ҵ�Ʊ��');window.close();</script>"
   Response.End
 end if
end if
%>