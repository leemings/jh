<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'С������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
'#####���䴦��#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=xhhz(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'С������
function xhhz(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select boysex,boy,��ż,hw FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
mypic="<img src=/chat/boy/"&rs("boysex")&" width=30 height=30>"
mypo=rs("��ż")
myhw=rs("hw")
cwnl="����"
rs.close
rs.open "select "& cwnl &",����,�ȼ�,����,grade,����,����ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ�����ձ�����ɱ�����ȷŹ����ɣ�');}</script>"
Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
Response.End
end if
if rs("�ȼ�")<=18 and rs("����")="��" and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
Response.End
end if
if rs("����")=True and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
Response.End
end if
rs.close
rs.open "select ����ʱ��,����,grade,����,ɱ���� from �û� where ����='"&aqjh_name&"'",conn,3,3
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ������������һ��ɣ�');}</script>"
Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
Response.End
end if
if rs("����")=True and aqjh_grade<10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ����!');}</script>"
Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) and aqjh_grade<10 then
conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
Response.Write "<script language=JavaScript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�����ٷ����ˣ�');}</script>"
rs.close
set rs=nothing 
conn.close
set conn=nothing
Response.End
end if
if isnull(myhw) or myhw="" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ������û��С������ȥ�����һ���ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
zt=split(myhw,"|")
if UBound(zt)<>12 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
if zt(0)<>fn1 then
 Response.Write "<script Language=Javascript>alert('��ʾ��С�����ֲ��ԣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
cssj=clng(DateDiff("d",zt(6),now()))
gs=clng(zt(4))
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
Response.Write "<script Language=Javascript>alert('��ʾ��С�����鲻�ã���������һ���ܣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
if clng(zt(2))<100 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��С���������������ܴ���ˣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
response.end
end if
rs.close
'����0|�Ա�1|����2|����3|����4|����5|����6|״̬7|ι������8|���ι��9|����10|ͷ��11|���鿴ʱ��12
cwsr=clng(DateDiff("d",zt(6),now()))
if cwsr<5 then
xhhz="<font color=green>��С��������<font color=" & saycolor & ">##��ʹ��С������%%������С������̫С�ˣ���������(˵����ͬ���ھ���������,С��������������5ȡ��������ȥ�����������������Է�������ֵ����������!)</font>"
exit function
end if
xhhzz=int(cwsr/5)
tgyjz1=int(tgyjz/xhhzz)
tgyjz2=tgyjz-tgyjz1
'conn.Execute ("update �û� set "&cwnl&"="&tgyjz1&" where ����='" & to1 &"'")
'conn.Execute ("update �û� set "&cwnl&"="&cwnl&"+"&tgyjz2&",cw='' where ����='" & aqjh_name &"'")
'conn.Execute ("update �û� set ����="&tgyjz1&" where ����='" & to1 &"'")
xhhz="<font color=green>��С��������<font color=" & saycolor & ">##<bgsound src=wav/tgyj.wav loop=1>ʹ��С��[<b>"&zt(0)&"</b>]����%%��ʹ%%�����½���:<font color=red>-"&tgyjz1&"</font>��,С����������(�˺�:"& xhhzz &"),ȫ���Ѫ�ⶼ����һ����������%%��ȥ��%%㲲�������������[<b>"&cwnl&"</b>]:-"&tgyjz2&"�㣬С����������ȡ�壬����ǰ������ź��þûص����������������</font>"
set rs=nothing
conn.close
set conn=nothing
end function
%>
