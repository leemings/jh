<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../config.asp"-->
<%'ץС͵
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<25 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵȼ�����25��,Ҳ��ץС͵��С��С͵����Ҳ��͵��!');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��ץС͵��</font><font color=" & saycolor & ">"+zuaxiaotou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function zuaxiaotou(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
fromdj=rs("�ȼ�")
if to1="���" or to1=aqjh_name or to1=Application("aqjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('ץС͵�����д��뿴��ϸ�ˣ���');}</script>"
	Response.End
exit function
end if
if rs("ְҵ")="С͵" then
zuaxiaotou="##:��û���ѽ�������С͵����ץʲôС͵ѽ��"
exit function
end if
rs.close
rs.open "SELECT �ȼ�,����,ְҵ FROM �û� WHERE ����='"&to1&"'",conn
todj=rs("�ȼ�")
if rs("����")="�ٸ�" then
   Response.Write "<script language=JavaScript>{alert('����û�и��ѽ���ٸ��Ļ���С͵!');}</script>"
   Response.End
end if
if rs("ְҵ")<>"С͵" then
zuaxiaotou="##:�㷢ʲô�񾭰���%%��û͵����������ץ����ʲô��"
exit function
end if
randomize timer
r=int(rnd*7)+1
Select Case r
Case 1
	conn.execute "update �û� set ְҵ='�ɱ�' where ����='"&to1&"'"
	conn.execute "update �û� set ����=����+200  where ����='"&aqjh_name&"'"
	zuaxiaotou="С͵[%%]��[##]�Ľ̵���,����Ͷ������,�Ĺ����²�����С͵![##]��������200��)"
Case 2
	if todj>fromdj then
		conn.execute "update �û� set ����=����-100  where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ְҵ='�ɱ�' where ����='"&to1&"'"
		zuaxiaotou="С͵[%%]�书��ǿ,[##]��������С͵�Ķ���,������С͵[%%]�����ˣ�[##]�����½�100��)"
	else
		zuaxiaotou="[##]�����סС͵[%%],��ϸһ��,ԭ�����Լ�����,���ǰ�С͵[%%]����!"
	end if 
Case 3
	conn.execute "update �û� set ����=����+100,����=����+100,����=����+10000,����=����-200 where ����='"&aqjh_name&"'"
	conn.execute "update �û� set ְҵ='�ɱ�',״̬='��' where ����='"&to1&"'"
	call boot(to1,"ץС͵�������ߣ�"&aqjh_name&","&fn1)
	zuaxiaotou="����һ����ս,�����[##]���ڽ�С͵[%%]�����鰸!(��ý���10000��,��������100��,��������100��,�����½�200��)"
Case 4
	conn.execute "update �û� set ����=����-100,����=����+10000 where ����='"&aqjh_name&"'"
	conn.execute "update �û� set ְҵ='�ɱ�',����=����-10000 where ����='"&to1&"'"
	zuaxiaotou="[##]����ץ[%%]��һ��ԭ����ʶ������[%%]��͵���������ָ�[##]10000����[##]Ҳ��Щ�Ķ�������͵�Ĳ����ҵĶ������ʹ��ڶ���ﳤ��ȥ..."
case else
	zuaxiaotou="[##]�����סС͵{%%},��ϸһ��,ԭ�����Լ�����,���ǰ�С͵{%%}����!"
End Select
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function 
%>