<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'���յ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<bgsound src=wav/go.wav loop=1><font color=green>�����յ��ӡ�<font color=" & saycolor & ">"+zsdz(mid(says,i+1))+"</font>"
call chatsay("�Զ�",towhoway,towho,saycolor,addwordcolor,addsays,says)

'�����յ��ӵ�
function zsdz(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,grade from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("grade")<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻������Ҳ���ǳ���,Ԫ��\n�������е��ӣ�');}</script>"
	Response.End
end if
mp=rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)+1
zsdz="[##]�������ҵ��˺���,��������["&mp&"]�˶�,������,���ڹ��е���!"&fn1&"��ӭ�ֵ��Ǽ���!<input  type=button value='����!' onClick=javascript:zsdz"&regjm&".disabled=1;window.open('../jhmp/mp1.asp?id="&mp&"','d') name=zsdz"&regjm&">"
end function
%>