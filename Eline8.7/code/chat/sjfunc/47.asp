<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_hy then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������Ҫ["&jhdj_hy&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>����顿<font color=" & saycolor & ">"+qiuhun(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���
function qiuhun(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�Ա�,��ż,����,����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sex=rs("�Ա�")
if rs("��ż")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������һ��һ���ƣ�������ʲô��');}</script>"
	Response.End
end if
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������300���˼ҿ�������ģ�');}</script>"
	Response.End
end if
if rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��û��5���Ǯ�����ܵǼǽ��ģ�');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,�Ա�,��ż,�ȼ� FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
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
if rs("��ż")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ʲô�������߲���ѽ��');}</script>"
	Response.End
end if
if rs("�ȼ�")<=5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ȡһ������5�����ˣ�û���Ӱɣ�');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����=����-50000 where ����='" & sjjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
qiuhun="<bgsound src=wav/mg.WAV loop=1>[##]��{%%}��飺<img src='img/29.gif'>"&fn1&"<input type=button value='��Ը��' onClick=javascript:tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('jiehun.asp?name=" & sjjh_name &"&yn=1&to1="&to1&"','d') name=tongyi"&regjm&"><input type=button value='��Ը��' onClick=javascript:;tongyi"&regjm+1&".disabled=1;tongyi"&regjm&".disabled=1;window.open('jiehun.asp?name=" & sjjh_name &"&yn=0&to1="&to1&"','d') name=tongyi"&regjm+1&">"
end function
%>
