<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->

<%'���鵶��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")

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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����鵶��<font color=" & sayscolor & ">"+jiuqingdao(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


function jiuqingdao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")

rs.open "select ����,w6,w9,����,�ȼ�,����ʱ�� FROM �û� WHERE  ����="&"'" & sjjh_name &"'",conn,3,3
fla=rs("����")
dj=rs("�ȼ�")
w6w=rs("w9")
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if fla<8000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��8000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<20 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ20��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,���� FROM �û� WHERE ����="&"'" & towho &"'",conn,2,2
money=int(rs("����")/3)
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('���ǹٸ���Ա��!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

wuyn=iswp(w6w,"���鵶")
if wuyn=0 then
	Response.Write "<script language=JavaScript>{alert('���о��鵶��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

conn.execute "update �û� set ����=����-8000,����ʱ��=now() where ����="&"'" & sjjh_name &"'"
fq=abate(w6w,"���鵶",1)
conn.execute "update �û� set  w9='"&fq&"' where ����='"&sjjh_name&"'"

r=int(rnd*2)+1
select case r
case 2
conn.execute "update �û� set ����=����-"&money&" where ����="&"'" & towho &"'"
jiuqingdao="##������γ�һ�Ѿ��鵶<bgsound src=wav/fx42.wav loop=1>��<img src='img/9.gif' width='44' height='44'>��ݺݵض���<font color=red>%%</font>������ȥ,�����ɣ�%%����һ��������������,�����׺���Ϊˮ��������ʧȥ1/3������"&money&"��......" 
rs.close
rs.open "select ���� from �û� where ����="&"'" & towho &"'",conn,2,2
if rs("����")<=0 then 
	conn.execute "update �û� set ״̬='��' where ����="&"'" & towho &"'"
	jiuqingdao="##������γ�һ�Ѿ��鵶<bgsound src=wav/ZR2199.wav loop=1>��<img src='img/9.gif' width='44' height='44'>��ݺݵض���<font color=red>%%</font>������ȥ,������,%%����һ��������������֧���������������ԹŶ������ް�..." 
	call boot(towho,"���鵶�������ߣ�"&sjjh_name&","&fn1)
end if
case else
jiuqingdao="##������γ�һ�Ѿ��鵶<bgsound src=wav/fx42.wav loop=1>��<img src='img/9.gif' width='44' height='44'>��ݺݵض���<font color=red>%%</font>������ȥ,������,�����ǰ�����ĺ�˭��û�̵���......"
end select
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>