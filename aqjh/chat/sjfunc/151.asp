<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���Ʊ�ʯ
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
says="<font color=green>��ħ��ʦ�����Ʊ�ʯ��<font color=" & sayscolor & ">"+peibashi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function peibashi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,�ȼ�,����,ְҵ FROM �û� WHERE ����="&"'" & aqjh_name &"'",conn,3,3
fali=rs("����")
w6w=rs("w6")
if rs("�ȼ�")<50 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ50��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

if fali<3000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��3000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"�챦ʯ")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('��û�к챦ʯ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"�̱�ʯ")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('��û���̱�ʯ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
wuyn=iswp(w6w,"����ʯ")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('��û������ʯ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select w6,w9 FROM �û� WHERE ����="&"'" & aqjh_name &"'",conn,3,3
fq=abate(rs("w6"),"����ʯ",1)
fq=abate(fq,"�챦ʯ",1)
fq=abate(fq,"�̱�ʯ",1)
conn.execute "update �û� set  w6='"&fq&"' where ����='"&aqjh_name&"'"

fq1=add(rs("w9"),"ħ����ʯ",1)
conn.execute "update �û� set  w9='"&fq1&"' where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����-3000 where ����="&"'" & aqjh_name &"'"
peibashi="##ȡ���졢�̡������ֱ�ʯ�����ֱ�ʯ�����һ��һ����â����<img src='img/look52.gif'>���ֱ�ʯ������ħ����ʯ." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>