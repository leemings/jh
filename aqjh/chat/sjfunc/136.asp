<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ս����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=green>����ս���š�</font><font color=" & saycolor & ">"+tzzm(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tzzm(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
fn1=trim(fn1)
mymp=rs("����")
mydj=rs("�ȼ�")
myzt=rs("����")+rs("����")+rs("�书")
if mymp="�ٸ�" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺���ң��ٸ���Ա��ֹ������');}</script>"
	Response.End
end if
if rs("���")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Լ������������Ű���');}</script>"
	Response.End
end if
if mydj<50 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺�ȼ���50����û�У�������������');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM �û� WHERE ����='" &fn1&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺���ʲô�����˲������ڣ�');}</script>"
	Response.End
end if
tomp=rs("����")
todj=rs("�ȼ�")
tozt=rs("����")+rs("����")+rs("�书")
if tomp="�ٸ�" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺���ң����ܶԹٸ���Ա������');}</script>"
	Response.End
end if
if rs("���")<>"����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺�Է��������Ű����㿴���˰ɣ�');}</script>"
	Response.End
end if
if mymp<>tomp then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺����ͬһ���ɽ�ֹ������');}</script>"
	Response.End
end if
if mydj<todj then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵȼ�̫���ˣ�ս�����Է���');}</script>"
	Response.End
end if
if myzt<tozt then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����״̬̫���ˣ���һ�������ɣ�');}</script>"
	Response.End
end if
tologintime=CDate(rs("lasttime"))
if tologintime>now()-5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���棺����5�����е�¼���㲻��������λ�ã�');}</script>"
	Response.End
end if
if rs("����")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��û��100�������Ӵ����ǲ��еģ�');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-1000000,grade=5,���='����' where ����='"&aqjh_name&"'"
conn.execute "update �û� set grade=1,���='����' where ����='"&fn1&"'"
Response.Write "<script language=JavaScript>{alert('��ϲ���Ѿ���["&fn1&"]���ж������֮λ��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
tzzm="##�书��ǿ����������["&fn1&"]�ı�������Ȼ�������ģ���������ٸ�����100������ŷѣ������ٳ����Ǻڰ����������˵����С���ǣ��Ժ�͸����һ��ˣ���֤���ǳ���ĺ����ģ���"
end function
%>