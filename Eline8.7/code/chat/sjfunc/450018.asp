<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���յ����wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")

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
'�����뿪������Ѩ�ж�
call dianzan(towho)
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
says="<font color=green>�����յ��⡿<font color=" & sayscolor & ">"+shendangao(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)




function shendangao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,����,�ȼ� FROM �û� WHERE ����="&"'" & sjjh_name &"'",conn,3,3
fla=rs("����")
dj=rs("�ȼ�")
w6w=rs("w6")
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��1000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<10 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ10��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-1000 where ����="&"'" & sjjh_name &"'"
fqsl=mywpsl(w6w,"���յ���")
if fqsl<5 then
Response.Write "<script language=JavaScript>{alert('����5�����յ�����');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"���յ���",5)
conn.execute "update �û� set  w6='"&fq&"' where ����='"&sjjh_name&"'"
conn.execute "update �û� set ����=����+500000 where ����="&"'" & sjjh_name &"'"
shendangao="##�ߺߵشӵ��ϼ���һ��ѩ���񵶰��Լ��Ķ������˿�������һ�鷨�����յ���<img src='pic/dz59.gif'>��ȥ���ٺ٣����ɣ�����ѽ���б��´����Ұ���<font color=red>##</font>����������500000��." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

end function 
%>