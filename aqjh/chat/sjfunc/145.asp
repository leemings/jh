<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�����ӵ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
%>
<!--#include file="pkif.asp"-->
<%
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
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�������ӵ���<font color=" & sayscolor & ">"+fashezi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function fashezi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn
fali=rs("����")
rs.open "select �ȼ�,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn
fali=rs("����")
if rs("����")="�ٸ�" and aqjh_grade<10 then
Response.Write "<script language=JavaScript>{alert('���ǹٸ���Ա��!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('�����ж�������Ķ���ÿСʱ��ǰ25���ӣ�����������ʱ�䣡');window.close();}</script>"
	Response.End 
end if
if rs("�ȼ�")<30 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ30��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,���� FROM �û� WHERE ����='" & towho &"'",conn
money=int(rs("����")/10)
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԸ߼�����Ա��վ���ط����Ա����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ�����Գ�������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w6 FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
w6w=rs("w6")
wuyn=iswp(w6w,"��ǹ")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('���о�����ǹ��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
wuyn=iswp(w6w,"�ӵ�")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('�����ӵ���');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if fali<8000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��8000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"�ӵ�",1)
conn.execute "update �û� set  w6='"&fq&"' where ����='"&aqjh_name&"'"

conn.execute "update �û� set ����=����-5000 where ����="&"'" & aqjh_name &"'"

randomize()

'rnd1=int(rnd()*3)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update �û� set ����=����-"&money&" where ����='" & towho &"'"
fashezi="##������γ�һ�Ѿ�����ǹ<bgsound src=wav/Phant012.wav loop=1>����ݺݵض���<font color=red>%%</font>����һǹ,%%����һ��������ʧȥ1/10������"&money&"��......" 
rs.close
rs.open "select ���� from �û� where ����='" & towho &"'",conn
if rs("����")<=-100 then 
conn.execute "update �û� set ״̬='��' where ����='" & towho &"'"
fashezi="##������γ�һ�Ѿ�����ǹ<bgsound src=wav/Phant012.wav loop=1>����ݺݵض���<font color=red>%%</font>����һǹ,%%����һ��������������֧�����������У��òҰ�..." 
  call boot(towho,"ǹɱ�������ߣ�"&aqjh_name&","&fn1)
end if
case else
fashezi="##������γ�һ�Ѿ�����ǹ<bgsound src=wav/Phant012.wav loop=1>����ݺݵض���<font color=red>%%</font>����һǹ,����ǹ��̫��û����......"
end select

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>