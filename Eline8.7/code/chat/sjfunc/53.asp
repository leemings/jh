<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���ˡ�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_qr then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����˹�����Ҫ["&jhdj_qr&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>�����ˡ�<font color=" & saycolor & ">"+qingren(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function qingren(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,��Ա��,�Ա�,��ż,����,����,����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sex=rs("�Ա�")
if rs("����")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������������ˣ���Ҫ�ȷ��֡�');}</script>"
	Response.End
end if
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������������300���˼ҿ�������ģ�');}</script>"
	Response.End
end if
if rs("��Ա��")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫ����һ��������Ҫ2Ԫ��Ա�𿨣�');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,�Ա�,��ż,�ȼ� FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
if rs("�Ա�")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����˰ɣ���������ͬ�����ģ�');}</script>"
	Response.End
end if
if rs("����")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���˼��������ˡ������Ӳ��ðɡ�');}</script>"
	Response.End
end if
if rs("�ȼ�")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����һ��5���ĵ����ˣ�û���ӣ�');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����=����-50000 where ����='" & sjjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.unLock
randomize()
regjm=int(rnd*3348998)
qingren="<bgsound src=wav/mg.WAV loop=1>[##]��{%%}����������˵�����<img src='img/ornament02.gif'>"&fn1&"<input  type=button value='����' onClick="&chr(34)&"javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('qingren.asp?name=" & sjjh_name &"&yn=1&to1="&to1&"','d')"&chr(34)&" name=tongyi"&regjm&"><input type=button value='����' onClick="&chr(34)&"javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('qingren.asp?name=" & sjjh_name &"&yn=0&to1="&to1&"','d')"&chr(34)&" name=tongyi"&regjm+1&">"
end function
%>