<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�ٱ���ͨ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>���ٱ���ͨ��<font color=" & sayscolor & ">"+banbianshu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function banbianshu(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,����,֪��,ְҵ FROM �û� WHERE ����="&"'" & aqjh_name &"'",conn,3,3
fla=rs("����")
if fla<2000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��2000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ְҵ")<>"����ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊ����ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-2000 where ����="&"'" & aqjh_name &"'"
randomize()
'r=int(rnd()*3)
r=int(rnd*7)+4
select case r
case 1
	zstemp=add(rs("w6"),"��ˮ",1)
	conn.execute "update �û� set ֪��=֪��+2,w6='"&zstemp&"' where ����='"&aqjh_name&"'"
	rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ����ˮ�����ķ���200�㣬֪������2��."
case 2
zstemp=add(rs("w6"),"��ʯ",1)
conn.execute "update �û� set ֪��=֪��+2,w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ����ʯ�����ķ���200�㣬֪������2��."
case 3
zstemp=add(rs("w6"),"������",1)
conn.execute "update �û� set ֪��=֪��+2,w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ�������㣬���ķ���200�㣬֪������2��."
case 4
zstemp=add(rs("w6"),"�����",1)
conn.execute "update �û� set ֪��=֪��+2,w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ������㣬���ķ���200�㣬֪������2��."
case 5
zstemp=add(rs("w6"),"С����",1)
conn.execute "update �û� set ֪��=֪��+10,w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ��С���㣬���ķ���200�㣬֪������10��."
case 6
zstemp=add(rs("w6"),"�ϻ���",1)
conn.execute "update �û� set ֪��=֪��+2,w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����һ���ϻ��⣬���ķ���200�㣬֪������2��."
case else
banbianshu="##��<bgsound src=wav/Phant008.wav loop=1>�������ǲ��ѽ��ʲô��û�����."
rs.close
end select
end function 
%>��<img src='picwords/1.gif'>��С���㣬���ķ���<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>��."
case 6
zstemp=add(rs("w6"),"�ϻ���",1)
conn.execute "update �û� set w6='"&zstemp&"' where ����="&"'" & aqjh_name &"'"
rs.close
banbianshu="##��ʩչ����<bgsound src=wav/Phant008.wav loop=1>�����<img src='picwords/1.gif'>���ϻ��⣬���ķ���<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>��."
case else
banbianshu="##��<bgsound src=wav/Phant008.wav loop=1>�������ǲ��ѽ��ʲô��û�����."
rs.close
end select
end function 
%>