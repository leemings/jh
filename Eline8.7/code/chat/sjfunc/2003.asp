<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'�������wWw.happyjh.com��
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
says="<font color=red>���������<font color=" & saycolor & ">"+djs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function djs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�ȼ�,ְҵ,����,grade,����,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*3)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*3)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
fn1=abs(fn1)
if fn1<100 then
Response.Write "<script language=JavaScript>{alert('Ҫ����������Ҳ������һ�ٰɣ���̫С����');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if sjjh_name=to1 or to1="���"  then
 djs=sjjh_name & "��Ա��˳�����?!�������һ��Լ�ʹ�õ����ѽ��"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("ְҵ")<>"ħ��ʦ" then
Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ʹ�õ����������ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("�ȼ�")<80 then
Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�õ����������Ҫ80����');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if fn1>50000 and rs("grade")<10 then
Response.Write "<script language=JavaScript>{alert('��������ܴ�������');}</script>"
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
if rs("����")<500 then
Response.Write "<script language=JavaScript>{alert('�㷨��������Ҫ500�ķ�����');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,��� FROM �û� WHERE ����='"&to1&"'",conn
if rs("���")>50000000 then
djs=to1&"[*^_^*]Ц���������ǲ��ԣ��������������ٸ��ɣ�лл!"
else
conn.execute "update �û� set ����=����-500,����=����+"&fn1/100&",����=����-" & fn1 & " where ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����+" & fn1*5 & " where ����='"&to1&"'"
djs=sjjh_name & "ħ��ʦ������ȡ��" & fn1 & "������ħ����һָ���͸���" & to1 &"��"& to1 &"�õ���"& fn1*5&"���׻���������,���˵�����˵лл��(��������"& fn1/100 &")"
end if

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>
