<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�ͻ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_jhdj<jhdj_hua then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ͻ���Ҫ["&jhdj_hua&"]���ſ��Բ�����');}</script>"
	Response.End
end if
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
says="<bgsound src=wav/faqian.wav loop=1><font color=red>���ͻ���Ϣ��<font color=" & saycolor & ">"+songhua(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�ͻ�
function songhua(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('��ʾ���������󣬸�ʽ���£�[��Ʒ��&����]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ʹ�����֣�');}</script>"
	Response.End 
end if
zswupin=trim(zt(0))
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>999 and instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��999��');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,w7 FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if  mywpsl(rs("w7"),zswupin)<wusl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&zswupin&"]����������("&wusl&")����');}</script>"
	Response.End
end if
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������������300���˼Ҳ�������Ļ���');}</script>"
	Response.End
end if
rs.close
rs.open "select i FROM b WHERE a='" & zswupin &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ʒ["&zswupin&"]��ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����');}</script>"
	Response.End
end if
sm=rs("i")
rs.close
set rs=nothing
conn.close
set conn=nothing
randomize()
regjm=int(rnd*3348998)
songhua="[##]��{%%}����<font color=red>"&wusl&"</font>��<img src='../hcjs/jhjs/images/"&sm&"'>" & zswupin &" ,Ҳ��֪���˼�Ը��Ը�����...<input type=button value='����' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('xianhua.asp?fromname="& aqjh_name &"&yn=1&toname="&to1&"&huaming="&zswupin&"&wpsl="&wusl&"','d') name=tongyi"&regjm&"><input type=button value='����' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('xianhua.asp?fromname="& aqjh_name &"&yn=0&toname="&to1&"&huaming="&zswupin&"&wpsl="&wusl&"','d') name=tongyi"&regjm+1&">"
end function
%>