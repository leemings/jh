<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���˽���
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
says="<font color=red>�����ͽ�����<font color=blue>"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,�ȼ�,ְҵ,����,����,ת�� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>�㷢ɵ��������Ұ���������������Ҫ����Ұ�����~������</font>"
	exit function
end if
if rs("�ȼ�")<300 and rs("ת��")<5 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>С�ӣ���ȼ�̫���ˣ����Թ˲�Ͼ�ˣ������ʸ�������ˣ�"
	exit function
end if
if fn1>10 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>##</font>�㷢ɵ��������Ұ���������������Ҫ����Ұ�����~������</font>"
	exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,����,����,�ȼ�,w5,��Ա�ȼ�,��Ա����,��¼ FROM �û� WHERE ����='" & to1 &"'" ,conn
mdate=rs("��¼")
if rs("��¼")=true then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    lbsw="<font color=blue>%%</font>���Ѿ��������˽�����Ŷ��<font color=blue>##</font>�㷢�����˰ɣ�лл֧��~"
	exit function
end if
if rs("�ȼ�")<30 then
lbsw=to1 & "�ĵȼ�̫����,����30�������ܽ�����"
else
zstemp=add(rs("w5"),"��Ǯ��",2)
conn.execute "update �û� set ��¼=true,w5='"&zstemp&"',���=���+100 where ����='" & to1 &"'"
lbsw="��ӭ<font color=blue>[%%]</font>�����齭�����ͣ����Ѿ�30�����ϣ�<font color=blue>[##]</font>��������ϵ��Ļ�ӭ�����䷢������Ǯ��<font color=red>2</font>�ţ����<font color=red>100</font>����лл֧�ְ��齭����չ�������������Ŷ��</font>"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>