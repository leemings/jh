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
r=int(rnd*17)+1
select case r
case 1
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>õ��</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower19.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"õ��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˽���MMҲ������ۻᡣ�յ�MM<font color=#0000FF>�¹�</font>����������[<b><font color=red>��õ��</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower20.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��õ��",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close

case 3
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower21.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>����������[<b><font color=red>��������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower22.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>����������[<b><font color=red>��ٺ�</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower23.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"��ٺ�",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 6
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>����������[<b><font color=red>������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower24.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������",9)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 7
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>����������[<b><font color=red>ӭ��Ц</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower25.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ӭ��Ц",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 8
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>�¼�</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower26.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"�¼�",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower27.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>Ũ���ƻ�</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower28.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"Ũ���ƻ�",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 11
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>ѩ����</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower29.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ѩ����",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 12
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower30.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 13
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>������˵</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower31.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������˵",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 14
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower32.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"������",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 15
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>ʾ��õ��</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower33.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"ʾ��õ��",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 16
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>����������</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower34.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����������",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 17
xunfaqi=aqjh_name & "��" & to1 & "�ļ������ͣ������˻����ϰ�Ҳ������ۻᡣ�յ����ĵ�<font color=#0000FF>�����ϰ�</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'><img src='picwords/9.gif'>��<img src='../hcjs/jhjs/images/flower35.gif'>��"
	rs.open "SELECT w7 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	xianhua=add(rs("w7"),"����",99)
	conn.execute "update �û� set  w7='"&xianhua&"' where ����='"&aqjh_name&"'"
	rs.close
		
		
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
