<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'Ѱ�ҿ�Ƭ
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
if f<1 or f>15 then
	Response.Write "<script language=JavaScript>{alert('���׵������ڻ�û�з��ſ�Ƭ��Ѱ�ҿ�Ƭʱ��ΪÿСʱ��ǰ15���ӣ�');window.close();}</script>"
	Response.End 
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��Ѱ�ҿ�Ƭ��<font color=" & saycolor & ">"+xunfaqi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ�ҿ�Ƭ
function xunfaqi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �ȼ�,��Ա�ȼ�,���,����ʱ��,ְҵ,ת�� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
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
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fla=rs("���")
if rs("���")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ�Ҳ����޷�ʩչѽ������Ҳ��80������');window.close();}</script>"
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
hydj=rs("��Ա�ȼ�")
if hydj<3 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ա�ȼ�Ϊ3���ſ��Խ���Ѱ�ҿ�Ƭ�������QQ865240608 ');}</script>"
	Response.End
end if

rs.close
conn.execute "update �û� set ���=���-50,����ʱ��=now()  where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*12)+1
select case r
case 1
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"����",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close


case 2
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��˰��</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"��˰��",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close

	
case 3
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>��Ǯ��</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"��Ǯ��",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 4
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"������",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 5
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>���׿�</font></b>]<img src='picwords/2.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"���׿�",2)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close

case 6
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>�ֻ���</font></b>]<img src='picwords/2.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"�ֻ���",2)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close

case 7
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"����",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close

case 8
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>������</font></b>]<img src='picwords/2.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"������",2)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 9
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"����",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close
	
case 10
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�յ��󷽵�<font color=#0000FF>���׵���</font>���͸�������[<b><font color=red>����</font></b>]<img src='picwords/1.gif'>�š�"
	rs.open "SELECT w5 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	kapian=add(rs("w5"),"����",1)
	conn.execute "update �û� set  w5='"&kapian&"' where ����='"&aqjh_name&"'"
	rs.close
case 11
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ����˴˻������һ�ѣ���֪�ս��žͲȵ�����㣬�书����200������˥����~"
	rs.open "SELECT �书 FROM �û� WHERE ����='"&aqjh_name&"'",conn
	conn.execute "update �û� set  �书=�书-200 where ����='"&aqjh_name&"'"
	rs.close
case 12
xunfaqi=aqjh_name & "����" & to1 & "�ļ������ͣ������˻��׵���Ҳ������ۻᡣ�Ͻ����ز����ܹ�ȥ����֪����Сʯͷ��ˤ��һ�ʣ���������200�㣬��������~��"
	rs.open "SELECT ���� FROM �û� WHERE ����='"&aqjh_name&"'",conn
	conn.execute "update �û� set  ����=����-200 where ����='"&aqjh_name&"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
