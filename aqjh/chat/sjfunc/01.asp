<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if zhuans<3 then
Response.Write "<script language=JavaScript>{alert('������Ҫ3��ת�����ϣ��������ʸ�ô?');window.close();}</script>"
    response.end
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
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if cwjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=" & saycolor & ">"+titl(mid(says,i+1))+"</font>"
call chatsay("������",towhoway,towho,saycolor,addwordcolor,addsays,says)
function titl(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����ͷ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
tx=rs("����ͷ��")
if rs("����")<10000  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ����Ҫ����10000���ú�������ɣ�",1)
end if
conn.execute "update �û� set ����=����-1000 where ����='" & aqjh_name &"'"
titl="<marquee width=100% behavior=alternate scrollamount=3><font color=AA00CC><img src="&tx &">[����]" & fn1 & "  (##)" & "</marquee>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>