<%@ LANGUAGE=VBScript.Encode%>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%'�͸�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_ti then
	Response.Write "<script language=JavaScript>{alert('��mtv������Ҫ"&jhdj_ti&"���ſ��Բ�����');}</script>"
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
says="<font color=green><img src=img/music.gif><font color=" & saycolor & ">"+mtv(fnn1,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function mtv(fn1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("����")<300 or rs("����")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������������300���˼Ҳ�������ĸ裡');}</script>"
	Response.End
end if
Set connt=Server.CreateObject("ADODB.CONNECTION")
connt.open Application("aqjh_usermdb")
sql="select * from swf where id="&fn1
Set Rs=connt.Execute(sql)
randomize()
regjm=int(rnd*3348998)
mtv="["&aqjh_name&"]��һ�ס�"&rs("����")&"��(MTV)�͸�{"&toname&"}���ͣ� <input type=button value='�տ�' onClick=javascript:shoukan"&regjm&".disabled=1;window.open('mtv/playswf.asp?id="&aqjh_name&"&name="&fn1&"&toname="&toname&"','playswf','scrollbars=no,resizable=no,width=500,height=400') name=shoukan"&regjm&" style='background-color:#86A231;color:FFFFFF;border: 1 double'>"
rs.close
set rs=nothing
Response.Write "<script language=JavaScript>parent.f2.af.mdsx.checked=true;parent.m.location.reload();</Script>"
end function
%>