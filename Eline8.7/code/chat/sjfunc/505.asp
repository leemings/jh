<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'��ҹ���wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_frag then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ҹ����Ҫ["&jhdj_frag&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>����ҹ�㡿<font color=" & saycolor & ">"+frag(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ҹ��
function frag(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ְҵ,����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if trim(towho)="" or towho="���" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
if rs("ְҵ")<>"��ҹ��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ְҵ���ǵ�ҹ�㣬���ܲ�����');}</script>"
	Response.End
end if

if rs("����")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��û��2���Ǯ�����ܵ�ҹ�㣡');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������������1000���ܵ�ҹ�㣡');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,����,����,�书 FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs("����")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է�û��2���Ǯ�����ܰ�����ҹ�㣡');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���������1000�����ܰ�����ҹ�㣡');}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_online")=to1
Application.UnLock
randomize()
regjm=int(rnd*3348998)
frag="��ҹ��ʦ<img src='pic/dz621.gif' width='40' height='55'>[##]��{%%}����⣺��һ�η���շ�2����������1000�<br><input style='font-size: 9pt; background-color: #99ccff; border-style: ridge' type=button value='���̽�����' onClick=javascript:;tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('frag.asp?fromname=" & sjjh_name &"&toname="&to1&"','d') name=tongyi"&regjm&">" 


 
end function
%>
