<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'û�շ�����wWw.happyjh.com��
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
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��û�շ�����<font color=" & saycolor & ">"+mushou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'û�շ���
function mushou(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,w10,�ȼ�,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<int(rnd*5) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=int(rnd*5)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
dj=rs("�ȼ�")
if rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��50000�㰡��');window.close();}</script>"
	response.end
end if
if dj<230 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ230��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select �ȼ�,w10 FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs("�ȼ�")<=10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���뽭���Ѿ������˾Ͳ�Ҫ������������');}</script>"
	Response.End
end if


duyao=rs("w10")
if isnull(duyao) or duyao="" then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�Է�ʲô������û�а���');</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,w10 from �û� where ����='"&sjjh_name&"'",conn,2,2
if rs("����")<2000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ʒ��Ҫ5000�������ѣ�');}</script>"
	Response.End
end if

dim js(10)
js(0) ="������"
js(1) ="����ʯ"
js(2) ="����ʯ"
js(3) ="����ʯ"
js(4) ="��ȡ��"
js(5) ="��������"
js(6) ="ʥ����"
js(7) ="��ڤ��"
js(8) ="����ȭ"
js(9)="������"
randomize()
myxy = Int(Rnd*10)

rs.close
rs.open "select w10 from �û� where ����='"&sjjh_name&"'",conn
duyao=abate(rs("w10"),js(myxy),1)
conn.execute "update �û� set w10='"&duyao&"' where  ����='"&to1&"'"
duyao=add(rs("w10"),js(myxy),1)
conn.execute "update �û� set w10='"&duyao&"',����=����-50000,����ʱ��=now() where  ����='"&sjjh_name&"'"

mushou=sjjh_name & "һ����ȣ�" & to1 & "˭�����ڷǶ���ʱ�����÷������˵ģ�û���������𣬳ͽ�һ���յ����"&js(myxy)&"һ����" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>


