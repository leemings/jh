<%
	if chatroomsn=1 or chatroomsn=2 or Application("yx8_mhjh_systemname"&chatroomsn)=mycorp or mycorp="�ٸ�"  then
	act=Trim(mid(msg,3,3))
	msg=Trim(mid(msg,6))
	if instr(msg,"'")<>0 or (msg="" and  instr(";����;��ʬ;����;����;����;��Ѩ;��Ѩ;�Զ�;����;����;����;�ر�;�Ӱ�;���;����;�ش�;�ϳ�;���;����;��ͽ;��ʦ;��ʦ;��ͽ;��λ;����;����;�鿴;����;��Ǯ;���;͵Ǯ;����;����;͵��;ն��;����;����;����;����;",";"&act&";")=0) then
		Response.Write "<script language=javascript>alert('"&act&"�����ĺ���Ӧ�ô��к��ʵ���䣬�磺//"&act&" ������');</script>"
		Response.End
	elseif instr(onlinename,";"&sendto&";")=0 and instr(";����;���;����;����;���;׷ɱ;��ԯ;����;��Ѩ;�Զ�;����;����;�Ӱ�;���;�ϳ�;��ͽ;��ʦ;��ʦ;��ͽ;��λ;����;����;����;�鿴;Ͷ��;��Ǯ;����;��λ;���;���;����;͵Ǯ;����;����;͵��;����;����;���;����;����;�̷�;����;ն��;����;����;",";"&act&";")<>0 and sendto<>"������" and sendto<>"ħ��ʦ" and sendto<>"������" then
		Response.Write "<script language=javascript>alert('"&act&"����˭����');</script>"
		Response.End
	elseif sendto=username and instr(";�ӽ�;����;���;����;����;���;׷ɱ;��ԯ;����;���;�Զ�;��λ;����;����;�Ӱ�;���;�ϳ�;����;����;��ͽ;��ʦ;��ʦ;��ͽ;��λ;����;����;����;Ͷ��;��Ǯ;����;���;���;����;͵Ǯ;����;����;͵��;����;�Ŵ�;����;�̷�;����;ն��;����;����;����;����;",act)<>0 then
		Response.Write "<script language=javascript>alert('"&act&"�����Լ�����');</script>"
		Response.End
	end if
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("yx8_mhjh_connstr")
	set rst=server.CreateObject("adodb.recordset")
	isprivacy=0
	select case act
case "�鿴"
msg=seeequip(username,sendto)
isprivacy=1
%><!--#include file="send/4.asp"--><%
case "����"
msg=gd(username,sendto,msg)
%><!--#include file="send/71.asp"--><%
case "����"
msg=huanying(username,sendto)
%><!--#include file="send/5.asp"--><%
case "����"
msg=nenyin(username,sendto,mygrade)
%><!--#include file="send/6.asp"--><%
case "����"
msg=chuanyan(username,sendto)
%><!--#include file="send/46.asp"--><%
case "�Ķ�"
msg=xindong(username,mygrade,msg)
%><!--#include file="send/47.asp"--><%
case "����"
msg=qiqi(username,sendto,msg)
%><!--#include file="send/9.asp"--><%
case "����"
msg=jmen(username,sendto,mycorp)
%><!--#include file="send/48.asp"--><%
case "����"
msg=hmen(username,sendto,mycorp)
%><!--#include file="send/49.asp"--><%
case "ŭ��"
msg=nuhou(username,sendto,mygrade,msg)
%><!--#include file="send/50.asp"--><%
case "����"
msg=xintiao(username,mygrade,msg)
%><!--#include file="send/51.asp"--><%
case "����"
msg=attack(username,sendto,msg,mygrade)
%><!--#include file="send/7.asp"--><%
case "Ͷ��"
msg=cast(username,sendto,msg)
%><!--#include file="send/8.asp"--><%
case "��Ǯ"
msg=sendmoney(username,sendto,msg)
%><!--#include file="send/14.asp"--><%
case "��ʦ"
msg=bais(username,sendto,mygrade)
%><!--#include file="send/15.asp"--><%
case "��ͽ"
msg=studi(username,sendto,mygrade)
%><!--#include file="send/16.asp"--><%
case "���"
msg=jiebai(username,sendto,mygrade)
%><!--#include file="send/17.asp"--><%
case "�Ӱ�"
msg=jiebaich(username,sendto)
%><!--#include file="send/18.asp"--><%
case "����"
msg=duanyi(username,sendto)
%><!--#include file="send/45.asp"--><%
case "���"
msg=bangfu(username,sendto,msg)
%><!--#include file="send/73.asp"--><%
case "��ʦ"
msg=tuoli(username,sendto)
%><!--#include file="send/20.asp"--><%
case "��ͽ"
msg=zhutu(username,sendto)
%><!--#include file="send/21.asp"--><%
case "����"
msg=dazuo(username,sendto)
%><!--#include file="send/22.asp"--><%
case "�Ŵ�"
msg=fangda(username,sendto,mygrade,msg)
%><!--#include file="send/23.asp"--><%
case "����"
msg=fengzhou(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/24.asp"--><%
case "�̷�"
msg=xingfa(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/26.asp"--><%
case "����"
msg=fakuan(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/27.asp"--><%
case "����"
msg=alert(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/28.asp"--><%
case "ն��"
msg=zhanshou(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/29.asp"--><%
case "��Ǯ"
msg=distributemoney(username,msg)
%><!--#include file="send/30.asp"--><%
case "���"
msg=degree(username,sendto,mycorp,msg)
%><!--#include file="send/32.asp"--><%
case "��λ"
msg=fengwei(username,sendto,msg)
%><!--#include file="send/53.asp"--><%
case "��λ"
msg=chuanwei(username,sendto,mycorp)
%><!--#include file="send/31.asp"--><%
case "���"
msg=dislodge(username,sendto,mycorp)
%><!--#include file="send/33.asp"--><%
case "����"
msg=impart(username,sendto,msg)
%><!--#include file="send/34.asp"--><%
case "͵Ǯ"
msg=touqian(username,sendto)
%><!--#include file="send/35.asp"--><%
case "����"
msg=burglemp(username,sendto)
%><!--#include file="send/36.asp"--><%
case "͵��"
msg=burglemoney(username,sendto)
%><!--#include file="send/37.asp"--><%
case "����"
msg=present(username,mygrade,sendto,msg)
%><!--#include file="send/38.asp"--><%
case "����"
msg=card(username,sendto,msg,mycorp)
%><!--#include file="send/39.asp"--><%
case "����"
msg=bulletin(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/40.asp"--><%
case "�Զ�"
msg=gedu1(username,sendto,msg)
%><!--#include file="send/57.asp"--><%
case "����"
msg=leave(username)
%>
<!--#include file="send/54.asp"--><%
case "����"
msg=comeback(username)
%>
<!--#include file="send/55.asp"--><%
case "����"
msg=shenlong(username,sendto,msg)
%>
<!--#include file="send/58.asp"--><%
case "��ԯ"
msg=xuanyuan(username,sendto,msg)
%>
<!--#include file="send/59.asp"--><%
case "׷ɱ"
msg=zuisha(username,sendto,mycorp,mygrade,msg)
%>
<!--#include file="send/61.asp"--><%
case "����"
msg=xw(username,sendto,msg)
%>
<!--#include file="send/72.asp"--><%
case "��ʬ"
msg=shoushi(username,mycorp,msg)
%>
<!--#include file="send/63.asp"--><%
case "����"
msg=zhangml(username,sendto,mycorp,msg)
%>
<!--#include file="send/65.asp"--><%
case "����"
msg=junzhul(username,sendto,msg)
%>
<!--#include file="send/66.asp"--><%
case "���"
msg=qiuhun(username,sendto,msg)
%>
<!--#include file="send/67.asp"--><%
case "����"
msg=jiangli(username,sendto,mycorp,msg)
%>
<!--#include file="send/68.asp"--><%
case "����"
msg=jijin(username,mycorp)
%>
<!--#include file="send/69.asp"--><%
case "�ӽ�"
msg=bujin(username,mycorp,msg)
%>
<!--#include file="send/70.asp"--><%
case else
			msg="<font color=FF0000>��ϵͳ��</font>##���Բ��𣬲�������޷�ִ��"	
	end select
	act=2
	set rst=nothing
	conn.Close
	set conn=nothing
	else
		Response.Write "<script language=javascript>alert('�㲻�Ǳ�����ӣ������ڴ���Ұ��');</script>"
		Response.End
	end if
%>