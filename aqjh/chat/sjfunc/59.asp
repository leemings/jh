<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���ɷ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('���ɷ�����Ҫ4���ſ��Բ�����');}</script>"
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
says=fakuan(mid(says,i+1)+0,towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���ɷ���"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
  sf=rs("����")
  menpai=rs("����")
if  sf<>"����" and sf<>"����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="[����]��[##]˵����û��Ȩ��ʹ�ñ����ܣ�"
    exit function
end if
rs.close
rs.open "SELECT ����,grade FROM �û� where ����='" & to1 & "' and ����='" & menpai & "'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[�ٷ�]��[##]˵��[%%]���������ɵ��˰���"
	exit function
end if
if aqjh_grade<=rs("grade") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[�ٷ�]��[##]˵��[%%]�����ɵ����Ż����ǳ��ϣ����ǲ��ܶ��������ģ�"
	exit function
end if

if fn1>1000000 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="##�뷣%%����" & fn1 & "����������Ŀ���������涨,����ʧ�ܣ������涨�����������Ϊ100��"
	exit function
end if
if rs("����")>=fn1 then
    conn.execute("update �û� set ����=����-" & fn1 & " where ����='" & to1 & "'")
    conn.execute "update p set h=h+" & fn1 & " where a='" & menpai & "'"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan="<font color=" & saycolor & "><font color=green>�������̷���</font>##��Ϊ</font></b>%%<font color=red><b>ȷʵ����ʲô���ˣ���������</b></font>" & fn1 & "<font color=red><b>���������ŵ����ɻ����У�</b></font>"
	exit function    
end if
fakuan="<font color=red>##</font></b><font color=#000000><b>��ͷŹ�</font></b><font color=red>%%</font><font color=#000000><b>�ɣ��������ܹ���û��" & fn1&"������!</b></font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>