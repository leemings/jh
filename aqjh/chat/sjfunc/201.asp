<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���ﲫ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="pkif.asp"-->
<%
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
 Session("aqjh_inthechat")="0" 
 Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=cwbd(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���﹥��
function cwbd(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
fn1=trim(fn1)
rs.open "select k,i FROM b WHERE a='" & fn1 &"'",conn
If Rs.Bof OR Rs.Eof then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ����ĳ�����ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����');}</script>"
 Response.End
end if
cwnl=rs("k")
mypic="<img src=../hcjs/jhjs/images/"&rs("i")&">"
rs.close
rs.open "select ����ʱ��,����,grade,����,ɱ����,cw from �û� where  ����='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ����ոձ�����ɱ������������һ��ɣ�');}</script>"
 Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
 Response.End
end if

if rs("grade")>=6 and rs("grade")<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
 Response.End
end if
if rs("����")=True and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ�����������������벻Ҫ����!');}</script>"
 Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
 conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
 Response.Write "<script language=javascript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�����ٷ����ˣ�');}</script>"
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.End
end if
if isnull(rs("cw")) or rs("cw")="" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ��û�г����ȥ����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ�������Ѿ��������������¹���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
cssj=clng(DateDiff("d",zt(2),now()))
gs=clng(zt(5))
if gs<(80+cssj*5) then
 czzt="����"
elseif gs>((100+cssj*5+50)) then
 czzt="˳��"
else
 czzt="��ͨ"
end if
if czzt="����" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ���������鲻�ã���������һ���ܣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if

rs.close
'����0|���1|����2|����3|����4|����5|�չ�����6|�չ�ʱ��7
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(4))/100))&"|"&(clng(zt(4))-int(clng(zt(4))/200))&"|"&zt(5)&"|"&zt(6)&"|"&now()
rs.open "select ����,�ȼ�,����,grade,����,����ʱ��,cw from �û� where ����='" & to1 &"'",conn
if isnull(rs("cw")) or rs("cw")="" then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ���Է���û�г���޷�������');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('��ʾ���Է������������޷�������');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
 response.end
end if
if rs("����")="����" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
 Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and aqjh_grade<10 then
 rs.close
 set rs=nothing 
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ�����ձ�����ɱ�����ȷŹ����ɣ�');}</script>"
 Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
 Response.End
end if
if rs("�ȼ�")<=18 and rs("����")="��" and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ���벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
 Response.End
end if
if rs("����")=True and aqjh_grade<10 then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script language=javascript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
 Response.End
end if
rs.close
temp1=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(4))/20))&"|"&(clng(zt(4))-int(clng(zt(4))/40))&"|"&zt(5)&"|"&zt(6)&"|"&now()
conn.Execute ("update �û� set cw='"&temp1&"' where ����='" & to1 &"'")
conn.Execute ("update �û� set cw='"&temp&"' where ����='" & aqjh_name &"'")
rs.open "select cw from �û� where ����='" & to1 &"'",conn
zt=split(rs("cw"),"|")
if clng(zt(3))<-100 then
 conn.Execute ("update �û� set cw='|' where ����='" & to1 &"'")
 conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','" & to1 & "�ĳ�����"&aqjh_name&"�������','����')"
 cwbd="<font color=green>�����ﲫ����<font color=" & saycolor & ">##ʹ���Լ��ĳ���"&mypic&"����%%�ĳ���[<b>"&zt(0)&"</b>]��ʹ[<b>"&zt(0)&"</b>]�����½�:<font color=red>-"&int(clng(zt(4))/20)&"</font>��,�����½�:<font color=red>-"&int(clng(zt(4))/40)&"</font>�㣬%%�ĳ�������������֧������������</font>"
else
 cwbd="<font color=green>�����ﲫ����<font color=" & saycolor & ">##ʹ���Լ��ĳ���"&mypic&"����%%�ĳ���[<b>"&zt(0)&"</b>]��ʹ[<b>"&zt(0)&"</b>]�����½�:<font color=red>-"&int(clng(zt(4))/20)&"</font>�㣬�����½�:<font color=red>-"&int(clng(zt(4))/40)&"</font>��</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>