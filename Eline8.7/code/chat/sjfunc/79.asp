<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�Ͷ��ӡ�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_sq then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Ͷ�����Ҫ["&jhdj_sq&"]���ſ��Բ�����');}</script>"
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

'�����뿪������Ѩ�ж�
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
says="<font color=green>���Ͷ��ӡ�<font color=" & saycolor & ">"+givedouzi(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�Ͷ���
function givedouzi(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>1000) and sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���Ͷ�������10�㣬���1000�㣡');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �ݶ�����,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("�ݶ�����")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô��Ķ����𣿿�����˵��');}</script>"
	Response.End
end if
if rs("����")<100000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ͷ�����Ҫ10�������ӣ�');}</script>"
	Response.End
end if
conn.execute "update �û� set �ݶ�����=�ݶ�����-" & fn1 & ",����=����-100000 where ����='" & sjjh_name &"'"
conn.execute "update �û� set �ݶ�����=�ݶ�����+" & fn1 & " where ����='" & to1 &"'"
givedouzi="<bgsound src=wav/thanks.WAV loop=1>##��" & fn1 & "�����͸���%%[$$redb"&fn1&"$$b]��ϣ�����ǿ��Գ�Ϊ�����ѣ��л���Ҳ����һЩ�ö���~~"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','����','�Ͷ���"&fn1&"��')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
