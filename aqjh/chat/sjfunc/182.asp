<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"�������" then
 Response.Write "<script language=javascript>{alert('��ʾ���ұ�����ȥ����������û�У�');}</script>"
 Response.End
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
says="<font color=red>��Ѱ���䱦��<font color=" & saycolor & ">"+bws(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function bws(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,ְҵ,����,grade,����,����,����ʱ��,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
if DateDiff("s",rs("����ʱ��"),now())<60 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('Ϊ�˷�ֹ����ˢ��Ѱ�ұ����30���Ӳ���!');}</script>"
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
if rs("�ȼ�")<80 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ñ�����������Ҫ80����');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<1000 then
	Response.Write "<script language=JavaScript>{alert('�����������1000�ˣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<10 then
	Response.Write "<script language=JavaScript>{alert('�����������10�ˣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<10000 then
	Response.Write "<script language=JavaScript>{alert('�������������Ҫһ������');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<1000 then
	Response.Write "<script language=JavaScript>{alert('�㷨��������Ҫ1000�ķ�����');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

dim js(11)
js(0) ="ҡǮ��"
js(1) ="���黷"
js(2) ="������"
js(3) ="˧����"
js(4) ="������"
js(5) ="�����"
js(6) ="�����"
js(7) ="�����"
js(8) ="��ȸ��"
js(9) ="������"
js(10) ="˧����"
randomize()
myxy = Int(Rnd*11)
rs.close
rs.open "select w6 from �û� where ����='"&aqjh_name&"'",conn
duyao=add(rs("w6"),js(myxy),1)
conn.execute "update �û� set w6='"&duyao&"',����=����-1000,����=����-10,����=����-10000,����=����-1000,����ʱ��=now() where  ����='"&aqjh_name&"'"
bws=aqjh_name & "��ʦ�����˰�������ڵõ���"&js(myxy)&"һ��ħ������" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>