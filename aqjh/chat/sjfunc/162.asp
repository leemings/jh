<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ħ����ʯ
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
says="<font color=green>��ħ����ʯ��<font color=" & sayscolor & ">"+molishi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function molishi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,����,�ȼ�,����,ְҵ FROM �û� WHERE ����="&"'" & aqjh_name &"'",conn,3,3
fla=rs("����")
dj=rs("�ȼ�")
w6w=rs("w9")
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
if rs("����")<2 then
Response.Write "<script language=JavaScript>{alert('������������޷�ʩչѽ������Ҳ��2�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<55 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ55��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

conn.execute "update �û� set ����=����-2000,����=����-2 where ����="&"'" & aqjh_name &"'"
wuyn=iswp(w6w,"ħ����ʯ")
	if wuyn=1 then
	fq=abate(w6w,"ħ����ʯ",1)
	conn.execute "update �û� set  w9='"&fq&"' where ����='"&aqjh_name&"'"
	else
	Response.Write "<script language=JavaScript>{alert('����ħ����ʯ��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
	end if

	conn.execute "update �û� set ����=����+10000,����=����+5 where ����="&"'" & aqjh_name &"'"
	molishi="##�ӿڴ����ó�ħ����ʯ������һ��,ħ����ʯ�����������̼�<font color=red>##</font>�ķ�������10000�㣬����ֵ����<font color=red>5</font>��." 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>icwords/0.gif'>��." 
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>