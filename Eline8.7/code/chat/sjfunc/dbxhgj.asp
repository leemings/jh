<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�ᱦС��������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####���䴦��#####
call roompd("�ᱦС������")
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
if towho=sjjh_name or towho="���" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Զ����ѻ�Ի����ˡ����ʹ�ã�');}</script>"
	Response.End
end if
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
says=xhgj(towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�ᱦС������
function xhgj(to1)
fn1=trim(fn1)
toname=to1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "select id,�ȼ�,����,grade,�Ա�,��ż,hw from �û� where ����='"&sjjh_name&"'",conn,3,3
if rs("����")=true and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����������������У����ܹ�����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("��ż")="" or rs("��ż")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û��飬�����ĺ��ӣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("�ȼ�")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��30�����Ϸ���ʹ��С��������');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if sjjh_grade>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ����ˣ������Թ�����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

mysex=rs("�Ա�")
mypo=rs("��ż")
myhw=rs("hw")
mynameid=rs("id")
rs.close
if mypo=to1 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����Զ����ѵ���ż���д˲�����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if mysex="Ů" then
	fyname=sjjh_name
else
	fyname=mypo
end if
if mysex="��" then
	rs.open "select hw from �û� where ����='"&mypo&"'",conn,3,3
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ���������Ҳ��������ż���ϣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
		response.end
	end if
	myhw=rs("hw")
	rs.close
end if
if isnull(myhw) or myhw="" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㻹û��С���أ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(myhw,"|")
if UBound(zt)<>12 then
	erase zt
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���������ݳ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

xhname=zt(0) 'С������
xhsex=zt(1) '�Ա�
xhtl=clng(zt(2)) '����
xhnl=clng(zt(3)) '����
xhgj=clng(zt(4)) '����
xhfy=clng(zt(5)) '����
xhday=zt(6) '����
xhzt=zt(7) '״̬
xhwy=clng(zt(8)) 'ι������
xhwyrq=zt(9) '���һ�ε�ι��ʱ��
xhgs=clng(zt(10)) '����
xhtx=zt(11)
xhck=zt(12)
erase zt
if xhzt="��" or xhtl<=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ĺ����Ѿ����ˣ�����ȥ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
cssj=clng(DateDiff("d",xhday,now()))
if cssj<=30 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('С������������30�죬���ܴ�ܣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if xhgs<(80+cssj*5) then
	czzt="�¶�"
elseif xhgs>((100+cssj*5+50)) then
	czzt="�Ҹ�"
else
	czzt="����"
end if
if czzt="�¶�" then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����С�����й¶����������һ���ܵģ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if xhtl<300 or xhnl<100 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����С�������������Ѳ����ٴ���ˣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.open "select id,����,�ȼ�,grade,����,����,�书,����,���� from �û� where ����='" & to1 &"'",conn,3,3
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
totl=rs("����")
tonameid=rs("id")
rs.close
tlkiller=int((xhgj+xhfy)/2)
wgkiller=int((xhtl+xhnl)/8)
xhtl=int(xhtl*19/20)
xhnl=int(xhnl*19/20)
xhgj=xhgj-5
xhfy=xhfy-5
xhgs=xhgs-10
temp=xhname&"|"&xhsex&"|"&xhtl&"|"&xhnl&"|"&xhgj&"|"&xhfy&"|"&xhday&"|"&xhzt&"|"&xhwy&"|"&xhwyrq&"|"&xhgs&"|"&xhtx&"|"&xhck
conn.execute "update �û� set ����=����-"&tlkiller&",�书=�书-"&wgkiller&" where ����='"&to1&"'"
conn.execute "update �û� set hw='"&temp&"' where ����='"&fyname&"'"
conn.execute "update �û� set �书=�书+"&wgkiller&" where ����='"&sjjh_name&"'"
newtotl=totl-tlkiller
xhgj="##�������ѵĺ���<img src='../../ico/"& xhtx &"-2.gif'><font color=red>"&xhname&"</font>����%%��"
fjh=""
if newtotl<-100 then
	conn.execute "update �û� set ״̬='��' where ����='"&to1&"'"
	conn.execute "update �û� set ɱ����=ɱ����+1 where ����='"&sjjh_name&"'"
	call boot(to1,"���У������ߣ�"&sjjh_name&",["&xhname&"]С����������")
	fjh="ɱ��������"&tlkiller&"�㣬��ȡ�书��"&wgkiller&"�㡣%%����һ����˫��һ�ţ����ڰ��޳�������ȥ��"
	bbb=zxrsbd(sjjh_name,mynameid)
	newfjh=xhgj&fjh
	newfjh=replace(newfjh,"##",sjjh_name)
	newfjh=replace(newfjh,"%%",toname)
	newfjh=newfjh&"<BR><BR>"&bbb
	fjh=fjh&"<BR><BR>"&bbb
	call dbxx(newfjh)
	to1=toname
else
	fjh="ɱ��������"&tlkiller&"�㣬��ȡ�书��"&wgkiller&"�㡣"
end if
xhgj=xhgj&fjh
set rs=nothing
conn.close
set conn=nothing
end function
%>