<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'�������ơ�wWw.51eline.com��
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
says="<font color=red>���������ơ�<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,�ȼ�,w10 FROM �û� WHERE ����='"&sjjh_name&"'",conn
dj=rs("�ȼ�")
fla=rs("����")
if rs("����")<20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��20000�㰡��');window.close();}</script>"
	response.end
end if
if dj<50 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ50��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w10")
if isnull(duyao) or duyao=abate(rs("w10"),"����ʯ",1) or duyao=abate(rs("w10"),"����ʯ",1) or duyao=abate(rs("w10"),"����ʯ",1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��������ʯȱһ���ɿ��������!');</script>"
	response.end
end if

rs.close
rs.open "select w10 from �û� where ����='"&sjjh_name&"'",conn
duyao=abate(rs("w10"),"����ʯ",1)
conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
duyao=abate(rs("w10"),"����ʯ",1)
conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
duyao=abate(rs("w10"),"����ʯ",1)
conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
duyao=add(rs("w10"),"ʥ����",1)
conn.execute "update �û� set  w10='"&duyao&"' where ����='"&sjjh_name&"'"
peibashi=sjjh_name & "<bgsound src=wav/m2.mid loop=1>ȡ���졢�̡���������ʯ��������ʯ�����һ��һ����â����<img src='img/look52.gif'>���ֱ�ʯ���ữ������ʥ����<img src='img/e5.jpg' width='150' height='80'><br></font>" & sjjh_name & "������ʥ����--���ɺ�������.�����ղ������־�ѧ��Ǭ����Ų��..����ȭ..������..�������Ƶȵ�..."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>