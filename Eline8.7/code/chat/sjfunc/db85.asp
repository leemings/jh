<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�ᱦ�����Ա�������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
call roompd("�ᱦ�����Ա�")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
towho=Trim(Request.Form("towho"))
if towho=sjjh_name or towho="���" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Զ����ѻ�Ի����ˡ����ʹ�ã�');}</script>"
	Response.End
end if
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
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

'�ᱦ�����Ա�����
function cwgj(fn1,to1)
fn1=trim(fn1)
toname=to1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select k,i FROM b WHERE a='" & fn1 &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĳ�����ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա��');}</script>"
	Response.End
end if
cwnl=rs("k")
mypic="<img src=../hcjs/jhjs/images/"&rs("i")&">"
rs.close
rs.open "select id,"& cwnl &",����,�ȼ�,����,grade,����,����ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
tgyjz=rs(cwnl)
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�����');}</script>"
	Response.End
end if
tonameid=rs("id")
rs.close
rs.open "select id,�ȼ�,����,grade,cw from �û� where  ����='"&sjjh_name&"'",conn,3,3
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й�����');}</script>"
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
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ��������¹���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
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
cwsr=clng(DateDiff("d",zt(2),now()))
if cwsr<5 then
	cwgj="<font color=green>���ᱦ�����Ա���<font color=" & saycolor & ">##��ʹ�ó����Ա�����%%�����Գ��ﻹ��̫����������(˵����ͬ���ھ���������,���������������5ȡ��������ȥ�����������������Է�������ֵ���������ӣ�)</font>"
	exit function
end if
cwgjz=int(cwsr/5)
tgyjz1=int(tgyjz/cwgjz)
tgyjz2=tgyjz-tgyjz1
conn.Execute ("update �û� set "&cwnl&"="&tgyjz1&" where ����='" & to1 &"'")
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"+"&tgyjz2&",cw='' where ����='" & sjjh_name &"'")
conn.Execute ("update �û� set ����="&tgyjz1&" where ����='" & to1 &"'")
cwgj="<font color=green>���ᱦ�����Ա���<font color=" & saycolor & ">##ʹ�ó���[<b>"&zt(0)&"</b>]"&mypic&"����%%��ʹ%%�����½���:<font color=red>-"&tgyjz1&"</font>��,�����ؼ�����(�˺�:"& cwgjz &"),%%������[<b>"&cwnl&"</b>]:-"&tgyjz2&"�㣬���ﻤ������ȡ�塭��</font>"
set rs=nothing
conn.close
set conn=nothing
end function
%>
