<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���⽱��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ�ô��⽱����');}</script>"
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=red>�����⽱����</font><font color=" & saycolor & ">"+zhouyu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���⽱��
function zhouyu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ֪��,����,����,�ȼ�,����,ְҵ,���,�书 FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
if rs("���")<50 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ50�����ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<5 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ5������ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<>"�ٸ�" then
	Response.Write "<script language=JavaScript>{alert('�㲻�ǹٸ���Ա������ʹ�ô˹��ܣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ���� FROM �û� WHERE ����='" & to1 &"'",conn
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԸ߼�����Ա��վ���ط����Ա����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-5 where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
zhouyu=to1 & "��ϲ��<img src=pic/h9.GIF>��������˱��������еĴ�������Ŀ������<font color=red>5</font>����ң���Ϊ���ϵͳ�������٣��Ǻǣ�лл������ף�´��и���!</font>"
 conn.execute "update �û� set ���=���+5 where ����='" & to1 &"'" 
case 2 
zhouyu=to1 & "��ϲ��<img src=pic/h13.GIF>��������˱��������еĴ�������Ŀ������<font color=red>6</font>����ң���Ϊ���ϵͳ�������٣��Ǻǣ�лл������ף�´��и���!</font>"
 conn.execute "update �û� set ���=���+6 where ����='"&to1&"'" 
case 3 
zhouyu=to1 & "��ϲ��<img src=pic/h21.GIF>��������˱��������еĴ�������Ŀ������<font color=red>10</font>����ң���Ϊ���ϵͳ�������٣��Ǻǣ�лл������ף�´��и���!</font>"
 conn.execute "update �û� set ���=���+10 where ����='"&to1&"'" 
case 4
zhouyu=to1 & "���ź�<img src=pic/h37.GIF>��������˱��������еĴ�������Ŀ��������<font color=red>1</font>����Ҷ��ѣ���Ϊ���ϵͳ��Ī��Ī�֣�лл������ף�´��и�������!</font>"
 conn.execute "update �û� set ���=���+1 where ����='"&to1&"'" 
case 5
zhouyu=to1 & "��ϲ��<img src=pic/h4.GIF>��������˱��������еĴ�������Ŀ������һ�Ƚ�<font color=red>1</font>�Ż�Ա�𿨣���Ϊ���ϵͳ�������٣��Ǻǣ�лл������ף�´��и���!</font>"
 conn.execute "update �û� set ��Ա��=��Ա��+1 where ����='"&to1&"'" 
case 6
zhouyu=to1 & "��ϲ��<img src=pic/h18.GIF>��������˱��������еĴ�������Ŀ�Ĵ󽱣�������ͷ�ȳ�����<font color=red>2</font>�Ż�Ա�𿨣�������������ã�ף�´�������˹��˹������!</font>"
 conn.execute "update �û� set ��Ա��=��Ա��+2 where ����='"&to1&"'"
   rs.close	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>