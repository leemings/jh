<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���˸���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
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
'�����뿪������Ѩ�ж�
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
says="<font color=red>�����˱���<font color=" & saycolor & ">"+xrfs()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xrfs()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if towho=Application("aqjh_automanname") or towho="���" then
 towho=aqjh_name
end if
if aqjh_name=towho then
'���Լ�
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if Session("xrfs")=true then
    Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ��б�������ˣ���Ҫ�����ˣ�');}</script>"
    Response.End
end if
if rs("times")<>3 then
    Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ�������״ε�½�����б��������');}</script>"
    Response.End
else
    conn.Execute ("update �û� set sl='����',slsj=now()+1,times=times+3 where ����='" & aqjh_name &"'")
    Session("xrfs")=true
end if
xrfs="[##]���������ֽ�����ͻȻ����һ����⣬���鸳�������ˣ�������������"
rs.close
set rs=nothing	
conn.close
set conn=nothing
else
'�Ա���
rs.open "select * FROM �û� WHERE ����='" & towho&"'",conn,2,2
if rs("times")<>3 then
    Response.Write "<script language=JavaScript>{alert('��ʾ��лл�㣬�����Ѿ�����Ҫ��������ˣ�');}</script>"
    Response.End
else
    conn.Execute ("update �û� set sl='����',slsj=now()+1,times=times+3 where ����='" & towho &"'")
    conn.Execute ("update �û� set ����=����+100 where ����='" & aqjh_name &"'")
end if
xrfs="[%%]���������ֽ�����[##]��[%%]ʹ�������˱���ͻȻ����һ����⣬���鸳��[%%]�����ˣ����һ�����֮�����ݵ��ٶ�2.5����[%%]������л��[##]��������100�㣡"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end if
end function
%>