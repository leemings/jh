<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ƽ����
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
says="<font color=red>��ƽ������<font color=" & saycolor & ">"+pas(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function pas(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,ְҵ,grade,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>50000 then
Response.Write "<script language=JavaScript>{alert('Ҫʹ��ƽ��������һ����಻�ܳ���5��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="���" then
 pas=aqjh_name & "��Ա��˳�����?!��������ʹ��ƽ����ѽ��"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
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
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ��ƽ����������Ҫ80����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if fn1>50000 and rs("grade")<10 then
	Response.Write "<script language=JavaScript>{alert('ƽ�������ܴ�������');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<fn1 then
	Response.Write "<script language=JavaScript>{alert('��������ô�������������ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<25000 then
	Response.Write "<script language=JavaScript>{alert('�㷨��������Ҫ500�ķ�����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�Ṧ,���� FROM �û� WHERE ����='"&to1&"'",conn
if rs("����")>300000 then
pas=to1&"Ц����������л" & aqjh_name & "�����������廹׳����,��л�����!�Ҳ���Ҫƽ����"
else
conn.execute "update �û� set ����=����-500,�Ṧ=�Ṧ+"&fn1/100&",����=����-" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����+"&fn1*10&" where ����='"&to1&"'"
pas=aqjh_name & "˫������һ�ӽ�" & fn1 & "������������" & to1 &"���ڣ�"& to1 &"������������������"& fn1*10&",���˵��۶���������[*^_^*]��(" & aqjh_name & "�Ṧ����"& fn1/100 &")"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>