<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'ħ����Ч��wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʹ��ħ����Ч��Ҫ[20]�����У�');}</script>"
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
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))

select case mycz
    case "/����ħ��"
        says="<font color=green>������ħ����</font><font color=" & saycolor & ">"+mofa1(mid(says,i+1),towho)+"</font>"
    case "/�ƶ�ħ��"
	says="<font color=green>���ƶ�ħ����</font><font color=" & saycolor & ">"+mofa2(mid(says,i+1),towho)+"</font>"
    case "/��ťħ��"
        says="<font color=green>����ťħ����</font><font color=" & saycolor & ">"+mofa3(mid(says,i+1),towho)+"</font>"
end select
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����ħ��
function mofa1(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ������|��ɫ|��С|���� \n\n    �磺�Ұ���|red|20|���� ');}</script>"
	Response.End 
end if

if not isnumeric(zt(2)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������[��С]��ʹ�����֣�');}</script>"
	Response.End 
end if
fontgb=trim(zt(0))
fontcolor=trim(zt(1))
fontsize=abs(int(clng(zt(2))))
fontt=trim(zt(3))
fontx=len(fontgb)*50
if fontsize>30 and fontsize<0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������С����С��0�Ҵ���30');}</script>"
	Response.End 
end if

Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if

if rs("����")<fontx then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ħ����Ҫ����"&fontx&"�㣡');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-1000,����=����-1000,����=����-"&fontx&" where ����='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa1="##ʹ��������ħ��(����������1000��,����1000��,����"&fontx&"��)��<br><font color="&fontcolor&" size="& fontsize&" face="&fontt&">"&fontgb&"</font>"
end function

'�ƶ�ħ��
function mofa2(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ������|��ɫ|����ɫ|left/right|�ٶ�|��ʱ \n\n    �磺�Ұ���|red|blue|right|20|500 ');}</script>"
	Response.End 
end if

if not isnumeric(zt(4)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������[�ٶ�]��ʹ�����֣�');}</script>"
	Response.End 
end if
if not isnumeric(zt(5)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������[��ʱ]��ʹ�����֣�');}</script>"
	Response.End 
end if

gdgb=trim(zt(0))
gdcolor=trim(zt(1))
gdbg=trim(zt(2))
gdtype=trim(zt(3))
gdspeed=abs(int(clng(zt(4))))
gdys=abs(int(clng(zt(5))))

if gdtype<>"left" and gdtype<>"right" then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ʽֻ�ж��֣�left��right');}</script>"
	Response.End 
end if

if gdspeed<0 or gdspeed>50 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ٶȴ���0��С��50');}</script>"
	Response.End 
end if
if gdys<100 or gdys>500 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʱ����100��С��500');}</script>"
	Response.End 
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ƶ�ħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ƶ�ħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if

if rs("����")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ƶ�ħ����Ҫ����500�㣡');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-1000,����=����-1000,����=����-500 where ����='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa2="##ʹ�����ƶ�ħ��(����������1000��,����1000��,����500��)��<br><font color="&gdcolor&"><marquee bgcolor="&gdbg&" direction="&gdtype&" scrollamount="&gdspeed&" scrolldelay="&gdys&">"&gdgb&"</marquee></font>"

end function

'��ťħ��
function mofa3(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ������|��ɫ|�������� \n\n    �磺�Ұ���|red|ϲ���� ');}</script>"
	Response.End 
end if


angb=trim(zt(0))
ancolor=trim(zt(1))
angb2=trim(zt(2))



Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ťħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ťħ����Ҫ����1000�㣡');}</script>"
	Response.End
end if

if rs("����")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ťħ����Ҫ����500�㣡');}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-1000,����=����-1000,����=����-500 where ����='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing

mofa3="##ʹ���˰�ťħ��(����������1000��,����1000��,����500��)��<br><input type=button value="&angb&" style=background-color:"&ancolor&" onclick=window.alert('"&angb2&"');>"

end function
%>