<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���콱��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_name=Session("aqjh_name")
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
says="<font color=red>�����˽�����<font color=" & saycolor & ">"+sljk()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function sljk()
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if Session("sljk_time")="" then Session("sljk_time")=rs("lasttime")
mycd=DateDiff("n",Session("sljk_time"),now())
if mycd<60 then
    if DateDiff("n",rs("lasttime"),now())>=30 then
       sljk="����һ��Сʱǰ�Ѿ�������ˣ�"
    else
       sljk="�㻹û���ݹ�һ��Сʱ�أ�"
    end if
    Response.Write "<script language=JavaScript>{alert('"&sljk&"');}</script>"
    Response.End
else
    conn.execute "update �û� set ��Ա��=��Ա��+2,mvalue=mvalue-2000 where ����='" & aqjh_name & "'"
    Session("sljk_time")=now()
end if
if rs("mvalue")<30000 then
Response.Write "<script language=JavaScript>{alert('���»���û��30000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�ȼ�")>100 then
Response.Write "<script language=JavaScript>{alert('��ȼ�����[100]������ʹ������Ŷ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ת��")>1 then
Response.Write "<script language=JavaScript>{alert('����ת���˲�������𿨣�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
sljk="���齭��վ������[##]�ݵ�����࣬�Ѿ���������<font color=red>"&mycd&"</font>�����ˣ��������<img src='picwords/2.gif'>������Ϊ����,�»��ּ���<img src='picwords/2.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>�㣡"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>nothing
end function
%>