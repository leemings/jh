<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���﹥��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Minute(time())>Application("sjjh_pktime") then
	call mess("��ʾ������PK���ʱ��ΪÿСʱǰ["& Application("sjjh_pktime") &"]�֣�����ʱ�䲻���Դ�ܣ�",1)
	response.end
end if
'#####���䴦��#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=cwgj(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���﹥��
function cwgj(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
rs.open "select k,i FROM b WHERE a='" & fn1 &"'",conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĳ�����ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����');}</script>"
	Response.End
end if
cwnl=rs("k")
mypic="<img src=../hcjs/jhjs/images/"&rs("i")&">"
rs.close
rs.open "select ����,�ȼ�,����,grade,����,����ʱ�� from �û� where ����='" & to1 &"'",conn
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ձ�����ɱ�����ȷŹ����ɣ�');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
if rs("�ȼ�")<=2 and rs("����")="��" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
if rs("����")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����ʱ��,����,grade,����,ɱ����,cw from �û� where  ����='"&sjjh_name&"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ������������һ��ɣ�');}</script>"
	Response.End
end if
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if

if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if
if rs("����")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ����!');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("sjjh_killman"))  and sjjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & sjjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("sjjh_killman") &"�����ٷ����ˣ�');}</script>"
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
	Response.Write "<script Language=Javascript>alert('��ʾ������û�г����ȥ����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("cw"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ����������¹���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
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
	Response.Write "<script Language=Javascript>alert('��ʾ���������鲻�ã���������һ���ܣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(3))<100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������������������ܴ���ˣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
'����0|���1|����2|����3|����4|����5|�չ�����6|�չ�ʱ��7
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-int(clng(zt(3))/10))&"|"&zt(4)&"|"&zt(5)&"|"&zt(6)&"|"&now()
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"-"&int(clng(zt(4))/10)&" where ����='" & to1 &"'")
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"+"&int(clng(zt(4))/10)&",cw='"&temp&"' where ����='" & sjjh_name &"'")
conn.Execute ("update �û� set ����=����-"&clng(zt(4))&" where ����='" & to1 &"'")
rs.open "select ���� from �û� where ����='" & to1 &"'",conn
if rs("����")<-100 then
	conn.Execute ("update �û� set ����ʱ��=now() where ����='" & to1 &"'")
	call boot(to1,"���У������ߣ�"&sjjh_name&",["&zt(0)&"]�ó������"&fn1)
	cwgj="<font color=green>�����﹥����<font color=" & saycolor & ">##ʹ�ó���[<b>"&zt(0)&"</b>]"&mypic&"����%%��ʹ%%�����½�:<font color=red>-"&clng(zt(4))&"</font>��,�����ؼ�����,%%������[<b>"&cwnl&"</b>]:-"&int(clng(zt(4))/10)&"��,%%����������֧���أ���["&zt(0)&"]ҧ��������䡭����</font>"
else
	cwgj="<font color=green>�����﹥����<font color=" & saycolor & ">##ʹ�ó���[<b>"&zt(0)&"</b>]"&mypic&"����%%��ʹ%%�����½�:<font color=red>-"&clng(zt(4))&"</font>��,�����ؼ�����,%%������[<b>"&cwnl&"</b>]:-"&int(clng(zt(4))/10)&"��</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>