<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'���֡�wWw.happyjh.com��
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
says="<font color=green>�����˷��֡�<font color=" & saycolor & ">"+fen(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���˷���
function fen(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ��Ա��,����,��� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
qingren=rs("����")
if rs("����")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ڻ�û�����˲��ܷ��֣�');}</script>"
	Response.End
end if
if rs("���")<2000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������200�򣬲���200����ַ�˭�ɣ�');}</script>"
	Response.End
end if
if rs("��Ա��")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫ�������˷�����Ҫ1Ԫ��Ա�𿨣�');}</script>"
	Response.End
end if
conn.execute "update �û� set ���=���-2000000,��Ա��=��Ա��-1,����='��' where ����='" & sjjh_name &"'"
conn.execute "update �û� set ���=���+2000000,����='��' where ����='"&qingren&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
fen="[##]������{"&qingren&"}200��ķ��ַѣ���{"&qingren&"}�����ˡ���"&fn1
end function
%>