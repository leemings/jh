<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ҡǮ��
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
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��ҡǮ����<font color=" & saycolor & ">"+ycs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function ycs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,w6,�ȼ�,����,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>3000000 then
Response.Write "<script language=JavaScript>{alert('ҡһ������100��಻�ܳ���3000000��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
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
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"ҡǮ��",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������һ��ҡǮ��Ҳû��Ү.���������!');</script>"
	response.end
end if

if rs("�ȼ�")<80 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ��ҡǮ���㻹��������Ҫ��80�����ϣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if fn1>3000000 then
	Response.Write "<script language=JavaScript>{alert('��̫̰�ˣ����𳬹�3����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<500000 then
	Response.Write "<script language=JavaScript>{alert('�㷨�������ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,���,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn
if rs("���")>400000000 then
ycs=aqjh_name & " �㻹��̰Ү������������̰��ѽ��...�ٺٺ�.(�����������������)��"
else
duyao=abate(rs("w6"),"ҡǮ��",1)
conn.execute "update �û� set w6='"&duyao&"',����=����-1000,���=���+"&fn1&" where ����='" & aqjh_name &"'"
ycs="��ϲ��ϲ,<font color=red size=2>" & aqjh_name & "</font>������ã���֪���������ҵ��˿�<font color=red size=2>ҡǮ��</font> ��ֻһҡ�͵���"&fn1&"�����ӡ�"

end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>