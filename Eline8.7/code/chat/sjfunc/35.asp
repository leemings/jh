<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(6)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]����û������¼�����������');}</script>"
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
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>������������<font color=" & saycolor & ">"+xiulian()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�������� xiulian
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,��������,�ȼ�,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")<>Application("sjjh_baowuname") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲢û�н�������["& Application("sjjh_baowuname") &"]��������');}</script>"
	Response.End
end if
if rs("��������")>=int(rs("�ȼ�")/Application("sjjh_baowuxl"))+1 then
	conn.execute "update �û� set ����='��',������=������+��������,������=������+��������,�书��=�书��+��������,����=����+("& Application("sjjh_baowuyin") &"*��������) where ����='" & sjjh_name &"'"
	xll=rs("��������")
	conn.execute "update �û� set ��������=0,����ʱ��=now() where ����='" & sjjh_name &"'"
	xiulian="##<bgsound src=wav/shuij.mid loop=1>ף����,������Ľ�������<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname") &"</font>�������<img src=sjfunc/18.gif>,�������޼�<font color=red>"& xll &"</font>��,���ֽ�+"& Application("sjjh_baowuyin")*xll &"��,����������ɣ��Զ���ʧ�ˡ�����������"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<(int(rnd*120)+1) then
	s=(int(rnd*120)+1)-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û��������ɣ����["& s &"]���ٽ��в�����');}</script>"
	Response.End
end if
if rs("����")<500  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������500���޷�������');}</script>"
	Response.End
end if
conn.execute "update �û� set ��������=��������+1,����=����-500,����ʱ��=now() where ����='" & sjjh_name &"'"
xiulian="<bgsound src=wav/xlsj.mid loop=1>##ӵ�н�������<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname") &"</font>����<img src=sjfunc/17.gif>�������������:"& rs("��������") & "�ν��б���������"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>