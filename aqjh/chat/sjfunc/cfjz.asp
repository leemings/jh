<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if application("aqjh_user")<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ����վ������Ȩ��������');}</script>"
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
says="<font color=green>����������<font color=" & saycolor & ">"+cfjz(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'������
function cfjz(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & to1 &"'",conn,2,2
guojia=rs("����")
if rs("�ȼ�")<10 then
  		Response.Write "<script language=JavaScript>{alert('��ʾ:�Է��ȼ�̫�ͣ�����ʤ�ξ�����ְλ��');}</script>"
		Response.End
end if
if rs("����")="�ٸ�" then
  		Response.Write "<script language=JavaScript>{alert('��ʾ:�ٸ����˲��ò��������ѡ��');}</script>"
		Response.End
end if
if Instr(Application("aqjh_admin_send"),"|" & to1 & "|")<>0 then 
  		Response.Write "<script language=JavaScript>{alert('��ʾ:�Է��ǲ��񰡣�');}</script>"
		Response.End
end if
conn.execute "update �û� set ְλ='����' where ����='" & to1 &"'"
cfjz="<bgsound src=wav/cfhk.mp3 loop=5><img src=pic/cfhk.gif>����վ������������%%Ϊ[" & guojia &"]�ľ�����##�䷢��Ʒ���������ף��ѽ��%%�ڽ���½������ʱ�����ܵ�����Ļ�ӭ!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','������')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>