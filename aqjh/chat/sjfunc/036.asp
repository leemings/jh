<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���ܽ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻����վ�����ܲ�����');}</script>"
	Response.End
end if
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if towho="���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ����������~��������~������');}</script>"
	Response.End
end if
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
says="<font color=red>�����˽�����<font color=blue>"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,�ȼ�,ְҵ,����,����,ת��,���� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>�����ƭ�ˣ�һ����<font color=red>" & fn1 & "</font>�˹���������������<font color=blue>[%%]</font>�ɣ�~�ֿ�ֻ��<font color=red>10</font>�������ˣ�</font>"
	exit function
end if
if rs("����")<>"�ٸ�" then
	Response.Write "<script language=JavaScript>{alert('���˽���ֻ�йٸ��Ĳſ��Բ������������ٸ�ô��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�ȼ�")<90 and rs("ת��")<1 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>��Ȼ���ǹٸ�������ȼ�Ҳ̫���˰�~`��90���ɣ�"
	exit function
end if
if fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>�����ƭ�ˣ�һ����<font color=red>" & fn1 & "</font>�˹���������������<font color=blue>[%%]</font>�ɣ�~�ֿ�ֻ��<font color=red>10</font>�������ˣ�</font>"
	exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,����,����,�ȼ�,w5 FROM �û� WHERE ����='" & to1 &"'" ,conn
if rs("�ȼ�")<30 then
lbsw=to1 & "�ĵȼ�̫���ˣ����ܱ�ʹ�ô˹��ܽ�����"
else
zstemp=add(rs("w5"),"�ӵ㿨",fn1)
conn.execute "update �û� set w5='"&zstemp&"',���=���+"&fn1*50&" where ����='" & to1 &"'"
lbsw="<font color=blue>[%%]</font>�������<font color=red>" & fn1 & "</font>λ���������ͣ�����<font color=blue>[##]</font>�˲�ȷ�Ϻ󣬰䷢�����ӵ㿨<font color=red>" & fn1 & "</font>�ţ����<font color=red>" &fn1*50& "</font>����лл֧�ֿ��ַ�չ�������������Ŷ��</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>