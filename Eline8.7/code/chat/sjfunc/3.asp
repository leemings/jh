<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��Ѩ��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_jx then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ѩ��Ҫ����ȼ�["&grade_jx&"]�ſ��Բ�����');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����Ѩ��</font><font color=" & saycolor & ">"+jie(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��Ѩ
function jie(fn1)
fn1=trim(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ״̬,��¼ from �û� where ����='" & fn1 & "'",conn
if rs("״̬")="��Ѩ" and (not rs.eof) and rs("��¼")>now() then
	conn.execute "update �û� set ��¼=now(),״̬='����',�¼�ԭ��='"&sjjh_name& ":������Ѩ!" &"' where ����='" & fn1 & "'"
	jie="##��" & fn1 & "ʹ���˽�Ѩ����" & fn1 & "��Ȼ�ѹ����ˡ���"
else
	jie="##��" & fn1 & "ʹ���˽�Ѩ��������" & fn1 & "��û�б���Ѩ����"
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& fn1 &"','��Ѩ','�����Ѩ')"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
