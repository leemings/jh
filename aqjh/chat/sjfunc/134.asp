<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��⻨��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��9���ſ��Բ�����');}</script>"
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
says="<font color=green>���ֻ�״Ԫ������<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��⻨��
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10 or fn1>1000) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ���ٵ���10����ң������ö���1000����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,����,���� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
hk=rs("����")
if hk>0 then 
Response.Write "<script language=JavaScript>{alert('�����ʲô������ѽ���Է��Ѿ����ֻ�״Ԫ�ˣ����Ҳ�������Ĺٸ���');}</script>"
	Response.End
else
conn.execute "update �û� set ���=���+" & fn1 & " where ����='" & to1 &"'"
conn.execute "update �û� set ����=20 where ����='" & to1 &"'"
end if
givegold="<bgsound src=wav/cfhk.mp3 loop=5><img src=pic/cfhk.gif>�������ȴ��һ����ͬ%%Ϊ�ֻ�״Ԫ��##�䷢��Ʒ����ƷΪ���["&fn1&"]��������ͻ���ʾף��ѽ��%%�ڽ���50�ε�½������ʱ�����ܵ�����Ļ�ӭ�������Ҳ�Ķ�Ҳ�����˼�%%һ���ǿ���������ͺ��ǻ�ȥ�ֻ�ȥ��!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','����ֻ�״Ԫ�������"&fn1&"��')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>