<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'Ѱ���ʻ�
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
says="<font color=red>��Ѱ���ʻ���<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ���ʻ�
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
dj=rs("�ȼ�")
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
f=Minute(time())
if f<25 or f>50 then
	Response.Write "<script language=JavaScript>{alert('�����ϰ�����û�п��Ż��⣬Ѱ���ʻ�ʱ��ΪÿСʱ��25-50���ӣ�');window.close();}</script>"
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
fla=rs("���")
if rs("���")<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ�Ҳ����޷�ʩչѽ������Ҳ��4������');window.close();}</script>"
	response.end
end if
if dj<100 and rs("ת��")<1 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ[100]��ս���ȼ�ѽ��ת���˳��⣩��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

rs.close
conn.execute "update �û� set ���=���-4,����ʱ��=now() where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*51)+1
select case r
case 1
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��֮��</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f4.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��֮��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 2
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>ȵ�����</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f3.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ȵ�����",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 3
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f21.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f5.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>õ������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f7.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"õ������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 6
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>���ް���</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f1.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"���ް���",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 7
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>������ζ</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f2.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��֮��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 8
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����İ�</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f8.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����İ�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f16.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��֮Ȫ</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f18.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��֮Ȫ",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f22.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 11
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>���˶�</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f19.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"���˶�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 12
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f14.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 13
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����˫��</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f13.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����˫��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 14
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f9.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 15
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��֮��</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f12.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��֮��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 16
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f24.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 17
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f6.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 18
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��Զ����</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f20.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��Զ����",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 19
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>���ж���</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f10.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"���ж���",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 20
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��Ľ</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f11.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��Ľ",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 21
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>һ������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/f4.jpg'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"һ������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 22
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>Ũ���ƻ�</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower17.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"Ũ���ƻ�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 23
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����ú�</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower18.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����ú�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 24
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>�����Ҹ�</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower19.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�����Ҹ�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 25
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>����ŨŨ</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower21.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����ŨŨ",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 26
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>������Զ</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower22.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������Զ",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 27
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>����������[<b><font color=red>һ������</font></b>]<img src='picwords/2.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower33.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"һ������",29)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 28
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower24.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 29
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����õ��</font></b>]<img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower25.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����õ��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 30
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>�ر�İ�</font></b>]<img src='picwords/3.gif'><img src='picwords/1.gif'>��<img src='../hcjs/jhjs/images/flower30.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�ر�İ�",31)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 31
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>�����õ��</font></b>]<img src='picwords/1.gif'><img src='picwords/1.gif'>��<img src='../hcjs/jhjs/images/flower390.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�����õ��",11)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 32
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>���ֵ�õ��</font></b>]<img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower350.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"���ֵ�õ��",8)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 33
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>ʧ����õ��</font></b>]<img src='picwords/7.gif'>��<img src='../hcjs/jhjs/images/flower320.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ʧ����õ��",7)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 34
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��ɵ�õ��</font></b>]<img src='picwords/6.gif'>��<img src='../hcjs/jhjs/images/flower400.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��ɵ�õ��",6)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 35
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>ҡ�ڵ�õ��</font></b>]<img src='picwords/5.gif'>��<img src='../hcjs/jhjs/images/flower330.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ҡ�ڵ�õ��",5)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 36
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>תȦ��õ��</font></b>]<img src='picwords/4.gif'>��<img src='../hcjs/jhjs/images/flower340.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"תȦ��õ��",4)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 37
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��л��õ��</font></b>]<img src='picwords/8.gif'><img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower360.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��л��õ��",88)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 38
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>ҡͷ��õ��</font></b>]<img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower370.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ҡͷ��õ��",8)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 39
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������õ��</font></b>]<img src='picwords/7.gif'>��<img src='../hcjs/jhjs/images/flower380.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������õ��",7)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 40
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��ͷ��õ��</font></b>]<img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower310.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��ͷ��õ��",8)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 41
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/1.gif'><img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower48.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",18)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 42
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������Զ</font></b>]<img src='picwords/1.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower49.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������Զ",19)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 43
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��ܰ�İ�</font></b>]<img src='picwords/1.gif'><img src='picwords/7.gif'>��<img src='../hcjs/jhjs/images/flower47.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��ܰ�İ�",17)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 44
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����ɴ�</font></b>]<img src='picwords/1.gif'><img src='picwords/6.gif'>��<img src='../hcjs/jhjs/images/flower45.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����ɴ�",16)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 45
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>ˮ��֮��</font></b>]<img src='picwords/2.gif'><img src='picwords/0.gif'>��<img src='../hcjs/jhjs/images/flower41.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ˮ��֮��",20)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 46
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������˵</font></b>]<img src='picwords/2.gif'><img src='picwords/1.gif'>��<img src='../hcjs/jhjs/images/flower40.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������˵",21)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 47
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>ǧ����˼</font></b>]<img src='picwords/2.gif'><img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower38.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ǧ����˼",28)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 48
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>���鲻��</font></b>]<img src='picwords/2.gif'><img src='picwords/6.gif'>��<img src='../hcjs/jhjs/images/flower36.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"���鲻��",26)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 49
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>�鶨һ��</font></b>]<img src='picwords/2.gif'><img src='picwords/7.gif'>��<img src='../hcjs/jhjs/images/flower34.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�鶨һ��",27)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 50
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>�ҵ�ǣ��</font></b>]<img src='picwords/1.gif'><img src='picwords/8.gif'>��<img src='../hcjs/jhjs/images/flower37.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�ҵ�ǣ��",18)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 51
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ�վ��<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��������</font></b>]<img src='picwords/3.gif'><img src='picwords/0.gif'>��<img src='../hcjs/jhjs/images/flower32.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",30)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
