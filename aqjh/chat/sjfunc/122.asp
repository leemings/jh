<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�ٸ�����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<8 then
	Response.Write "<script language=JavaScript>{alert('����ʲôѽ����Ĺ���ȼ��ɲ���ѽ��');}</script>"
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>���ٸ�������<font color=" & saycolor & ">"+jianglila(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function jianglila(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,grade,��� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<>"�ٸ�" then
      Response.Write "<script language=JavaScript>{alert('�㲻�ǹٸ���Ա���뵷����');}</script>"
end if
fn1=int(abs(fn1))
if fn1>5000000 or fn1<1000  then
Response.Write "<script language=JavaScript>{alert('�ٸ����������Գ���500���ٲ�������1000�ģ�');}</script>"
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        Response.End
end if
conn.execute "update �û� set ����=����+" & fn1 & " where ����='" & to1 & "'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','�ٸ���������"&fn1&"��')"
jianglila="��("& to1 &")�Խ����й��ף�"&aqjh_name & "�ѹٸ������Ľ��𷢸�(" & to1 &")"& fn1 & "����"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
