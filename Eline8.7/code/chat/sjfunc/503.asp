<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'��ǩ��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_chi then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ǩ��Ҫ["&jhdj_chi&"]���ſ��Բ�����');}</script>"
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
if towho=sjjh_name or towho=application("sjjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����ǩ��<font color=" & saycolor & ">"+chice(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ǩ
function chice(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,ְҵ,����ʱ��,����,����,����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������300���˼Ҳ���Ҫ������ǩ����');}</script>"
	Response.End
end if
if rs("����")<200000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����������ʱ���Ը�Ҳ�����Ǯ�ɣ������������ǩ�����˼ұ��㣡');}</script>"
	Response.End
end if
if rs("ְҵ")<>"����ʦ" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������ְҵ��������ʦ������������������Ϊ���˽�����������');</script>"
	response.end
end if
if rs("����")<800 or rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������Ҫ�ķ�800������1000��������̫���˰ɣ���');</script>"
	response.end
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('������ǩ60��һ�Σ����"& s &"��,�ɱ����ţ�');</script>"
	response.end
end if
rs.close
if to1<>"���" then
	rs.open "select ���� FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
	if rs("����")<200000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ���Է�������20��鶼û�У���������ʲô��ѽ��');}</script>"
		Response.End
	end if
	rs.close
end if
conn.execute "update �û� set ����ʱ��=now() where ����='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application("sjjh_smxs")=sjjh_name&"|"&to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
if to1<>"���" then
	chice="������ʦ<img src='../picc/05.gif' width='40' height='55'>[##]��{%%}�����ʲ�������������������Σ���һ֧ǩ�ɣ��۸񹫵����շѺ���һ��5��������׼��ҪǮ��<img src='../picc/chice.gif' width='60' height='50'><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='�˳�' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=�˳�','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='��Ե' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=��Ե','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='��ҵ' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=��ҵ','d') name=shiye"&regjm&">"
else
	chice="������ʦ<img src='../picc/05.gif' width='40' height='55'>[##]���������к�����˭Ҫ��ǩѽ���۸񹫵����շѺ���һ��5��������׼��ҪǮ��<img src='../picc/chice.gif' width='60' height='50'><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='�˳�' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=�˳�','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='��Ե' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=��Ե','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='��ҵ' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('chice.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=��ҵ','d') name=shiye"&regjm&">"	
end if
end function
%>