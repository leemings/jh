<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'Ѱ�һ�Ա��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
f=Minute(time())
if f<25 or f>40 then
	Response.Write "<script language=JavaScript>{alert('վ������û�з��Ż�Ա�𿨣�Ѱ�ҽ�ʱ��ΪÿСʱ��25-40���ӣ�');window.close();}</script>"
	Response.End 
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��Ѱ�һ�Ա�𿨡�<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ�һ�Ա��
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,��Ա�ȼ�,����ʱ��,ת��,ְҵ,�ȼ� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 then
	ss=3-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"������Ѱ�ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
dj=rs("�ȼ�")
if dj<120 and rs("ת��")<1 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ[120]��ս���ȼ�ѽ����ת���˳��⣩');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fla=rs("���")
if rs("���")<6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ�Ҳ����޷�ʩչѽ������Ҳ��6������');window.close();}</script>"
	response.end
end if
hydj=rs("��Ա�ȼ�")
if hydj<2 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ա�ȼ�Ϊ2���ſ��Խ���Ѱ�һ�Ա��!');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ���=���-6,����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=aqjh_name & "��Ѱ���˽����������������ҵ���[<b><font color=red><img src='picwords/1.gif'></font></b>]�Ż�Ա��<img src='card/jk.gif'>," & aqjh_name & "��������ѽ!" 
	conn.execute "update �û� set ��Ա��=��Ա��+1 where ����='" & aqjh_name &"'"
	

case 2
xunshuijing=aqjh_name & "��Ѱ���˽����������������ҵ���[<b><font color=red><img src='picwords/2.gif'></font></b>]�Ż�Ա��<img src='card/jk.gif'>," & aqjh_name & "��������ѽ!" 
	conn.execute "update �û� set ��Ա��=��Ա��+2 where ����='" & aqjh_name &"'"
	
case 3
xunshuijing=aqjh_name & "�����ˣ���Ѱ���˽������������ҵ���[<b><font color=red><img src='picwords/5.gif'></font></b>]�����" & aqjh_name & "�����˽��[<b><font color=red><img src='picwords/5.gif'></font></b>]��<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update �û� set ���=���+5 where ����='" & aqjh_name &"'"

case 4
xunshuijing=aqjh_name & "�����ˣ���Ѱ���˽������������ҵ���[<b><font color=red><img src='picwords/5.gif'></font></b>]�����" & dljh_name & "�����˽��[<b><font color=red><img src='picwords/5.gif'></font></b>]��<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update �û� set ���=���+5 where ����='" & aqjh_name &"'"

case 5
xunshuijing=aqjh_name & "�ÿ�ϧ����Ѱ���˽������������ҵ���[<b><font color=red><img src='picwords/5.gif'></font></b>]�����" & dljh_name & "�����˽��[<b><font color=red><img src='picwords/5.gif'></font></b>]��<img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'><img src='img/jinbi.gif'>!" 
	conn.execute "update �û� set ���=���+5 where ����='" & aqjh_name &"'"

case 6
xunshuijing=aqjh_name & "��Ѱ���˽����������������ҵ���[<b><font color=red><img src='picwords/5.gif'></font></b>]�Ż�Ա��<img src='card/jk.gif'>," & aqjh_name & "���[<b><font color=red><img src='picwords/3.gif'><img src='picwords/0.gif'></font></b>]�����!" 
	conn.execute "update �û� set ���=���-30,��Ա��=��Ա��+5 where ����='" & aqjh_name &"'"
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
