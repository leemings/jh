<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>���Ͻ��«��<font color=" & saycolor & ">"+xiulian()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�������� xiulian
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select * from �ᱦ",conn
slid=rs("�ھ�")
xlcs=rs("��������")
xb=rs("����")
lq=rs("��ȡ")
rs.close
if xb=false then	'���δ������������
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�ᱦ���������δ�������Ͻ��«������ģ�');}</script>"
	Response.End
end if
if lq=true then	'�������ȡ������������
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Ͻ��«���Ѿ��������ˣ��㻹������ʲô����');}</script>"
	Response.End
end if

rs.open "select id,�ȼ�,����,����,����,����ʱ�� FROM �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("id")<>slid then	'����ǹھ���������
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�Ƕᱦ�����Ĺھ����Ͻ��«������ģ�');}</script>"
	Response.End
end if

cs=int(rs("�ȼ�")/40)+1	'��������Ҫ�Ĵ���
if xlcs>=cs then	'��������������ˣ���������
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���Ѿ��������Ͻ��«�ˣ�');}</script>"
	Response.End
end if
sxyl=(xlcs+1)*10000000
if rs("����")<sxyl then	'����ʱ��������
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ǵ�"&xlcs+1&"�����������Ҫ"&sxyl&"���������Ǯ������');}</script>"
	Response.End
end if
if cs<=(xlcs+1) then	'�Ƿ�Ϊ���һ������
	conn.execute "update �ᱦ set ��������="& cs &",��ȡ=true"
	conn.execute "update �û� set ����=����-"& sxyl &",���=���+500,allvalue=allvalue+3000,������=������+1000,������=������+1000,�书��=�书��+1000,����ʱ��=now() where ����='" & sjjh_name &"'"
	xiulian="<bgsound src=wav/xlw.mid loop=1>##ף�������ᱦ������Ʒ��<font color=red><b>�Ͻ��«</b></font>�������<img src=sjfunc/18.gif>�����<font color=red><b>�������������书���޸���1000�㣬�õ�500����ң���������3000��</b></font>����֪�½�ᱦ�����ھ������ǡ�����"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<180 then
	s=180-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û��������ɣ����["& s &"]���ٽ��в�����');}</script>"
	Response.End
end if
if rs("����")<5000 or rs("����")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������������5000���޷�������');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-5000,����=����-5000,����=����-"&sxyl&",����ʱ��=now() where ����='" & sjjh_name &"'"
conn.execute "update �ᱦ set ��������=��������+1"
xiulian="<bgsound src=wav/xlhl.mid loop=1>##��ý����ᱦ�����ھ����õ���������<font color=red><b>�Ͻ��«</b></font>���Ǻǣ��б������Ǻã��������<font color=#000000><b>"& xlcs+1 & "</b></font>�ν��б���<img src=sjfunc/xlhl.gif>������"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>