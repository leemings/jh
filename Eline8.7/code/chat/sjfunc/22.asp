<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'���ˡ�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_grade<grade_tr then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����˹�����Ҫ["&grade_tr&"]���Ĳſ��Եģ�');}</script>"
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
says="<font color=#ff5600>�����ˡ�<font color=" & saycolor & ">"+tiren(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function tiren(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
fn1=Replace(fn1,"=","")
fn1=Replace(fn1,"untion","")
fn1=Replace(fn1,chr(39),"")
rs.open "select grade FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("grade")<sjjh_grade then
	tiren="<bgsound src=wav/gf.wav loop=1>һ�����%%�γ���������(�Ĺ�)"
	call boot(to1,"���ˣ������ߣ�"&sjjh_name&","&fn1)
else
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Ϊ���Ǹ߼�����Ա���㲻��������');}</script>"
	Response.End
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','����','"& fn1 &"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>