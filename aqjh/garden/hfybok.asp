<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=trim(request("myid"))
if name="" then
 Response.Write "<script Language=JavaScript>alert('��ʾ������!');window.close();</script>"
 Response.End
end if
if name<>aqjh_name then
 Response.Write "<script Language=JavaScript>alert('��ʾ����ʲô�Ұ����ֲ������н�!');window.close();</script>"
 Response.End
end if
%>
<!--#include file="data.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from ��԰����",connt,3,3
jz=rs("ͶƱ����")
if cdate(jz)>now() then
 Response.Write "<script Language=JavaScript>alert('��ʾ����û�п����أ��㼱ʲôѽ��');window.close();</script>"
 Response.End
end if
rs.close
sql="Select * from ��԰���ư� where ������='"&name&"'"
rs.open sql,connt,1,1
if rs.recordcount=0 then
 Response.Write "<script Language=JavaScript>alert('��ʾ�����˲�û�в�������');window.close();</script>"
 Response.End
else
   hjsj=rs("��ʱ��")
   if rs("�콱")=true then
 Response.Write "<script Language=JavaScript>alert('��ʾ�����ʲô�������Ѿ�������ˣ�');window.close();</script>"
 Response.End
   end if
   if rs("����")<1 or rs("����")>3 then
 Response.Write "<script Language=JavaScript>alert('��ʾ������û�н�������ʲô����');window.close();</script>"
 Response.End
   end if
   sj=DateDiff("d",hjsj,now())
   if sj>15 then
 Response.Write "<script Language=JavaScript>alert('��ʾ����Ľ�Ʒ�Ѿ����ڣ�');window.close();</script>"
 Response.End
   end if
  select case rs("����")
  case 1
    sql="update �û� set ����=����+100,���=���+500,��Ա��=��Ա��+10 where ����='"&name&"'"
    mess="�ٻ�԰����<font color=red>һ�Ƚ�</font>���ٸ��������500�����ֽ�(��Ա��)10Ԫ��վ�����԰佱����ϲ�ˣ�"
  case 2
    sql="update �û� set ����=����+50,���=���+300,��Ա��=��Ա��+5 where ����='"&name&"'"
    mess="�ٻ�԰����<font color=red>���Ƚ�</font>���ٸ��������300�����ֽ�(��Ա��)5Ԫ��վ�����԰佱����ϲ�ˣ�"
  case 3
    sql="update �û� set ����=����+20,���=���+100,��Ա��=��Ա��+3 where ����='"&name&"'"
    mess="�ٻ�԰����<font color=red>���Ƚ�</font>���ٸ��������100�����ֽ�(��Ա��)3Ԫ��վ�����԰佱����ϲ�ˣ�"
  case else
     Response.Write "<script Language=JavaScript>alert('��ʾ����Ʒ����');window.close();</script>"
     Response.End
  end select
  sql2="update ��԰���ư� set �콱=true where ������='"&name&"'"
  conn.execute(sql)
  connt.execute(sql2)
end if
rs.close
sql="Select hua from �û� where ����='"&aqjh_name&"'"
rs.open sql,conn,1,1
myhua=trim(rs("hua"))
fhua=split(myhua,"|")
huart="0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|0;0;0|"
if myhua<>"" then
seedsm=int(fhua(0))
if seedsm>0 then
huasj=fhua(1)
myhua_ok=seedsm&"|"&huasj&"|"&huart
else
myhua_ok="3|"&now()&"|"&huart
end if
else
myhua_ok="3|"&now()&"|"&huart
end if
conn.execute "update �û� set hua='"&myhua_ok&"' where ����='"&aqjh_name&"'"
call showchat ("<bgsound src=../chat/wav/hkjr.mp3 loop=2><font color=red>���ֻ�������</font><img src=pic/hk.gif>"&aqjh_name&mess)
Response.Write "<script Language=JavaScript>alert('��ʾ����ϲ���Ѿ���ȡ�˻�԰������Ʒ');window.close();</script>"
Response.End
%>