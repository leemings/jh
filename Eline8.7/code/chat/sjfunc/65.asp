<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'�Ĳ���wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�ܹ��Ĳ���');}</script>"
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��˫�˶Ĳ���<font color=" & saycolor & ">"+grdb(mid(says,i+1),towho)+"</font></font>"
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function grdb(fn1,toname)
if fn1<>1 and fn1<>2 and fn1<>3 then
	Response.Write "<script language=JavaScript>{alert('�������,Ӧ����1-3�����֣�');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ��ż,����,����,���� FROM �û� WHERE ����='"&sjjh_name&"'",conn,2,2
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������������300���˼Ҳ�����ģ�');}</script>"
	Response.End
end if
if rs("����")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('û��10���Ǯ�����ܶĲ���');}</script>"
	Response.End
end if
rs.close
rs.open "select ���� FROM �û� WHERE ����='"&toname&"'",conn,2,2
if rs("����")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��û��10���Ǯ��');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����=����-100000 where ����='"&sjjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_dubos")=toname
Application.unLock
randomize()
regjm=int(rnd*3348998)
grdb="["&sjjh_name&"]��{"&toname&"}����Ĳ�Ҫ�󣬶�ע10��׻��������ӣ�<input type=button value='ʯͷ' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=1&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=sitou"&regjm&"><input type=button value='����' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=2&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=jianzi"&regjm&"><input type=button value='��' onClick="&chr(34)&"javascript:sitou"&regjm&".disabled=1;jianzi"&regjm&".disabled=1;bu"&regjm&".disabled=1;window.open('dubo.asp?name="&sjjh_name&"&db=3&sjjh_bingwen="&fn1&"','d')"&chr(34)&" name=bu"&regjm&">"
end function
%>