<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('��Ҫ20�ȼ���');}</script>"
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
says="<font color=green>�����տ��֡�<font color=" & saycolor & ">"+lenyin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lenyin(fn1,to1)
if to1=aqjh_name or to1="���" then
   Response.Write "<script language=JavaScript>{alert('���ն�˭��');}</script>"
   response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,��Ա�ȼ� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<20000 then
  Response.Write "<script language=JavaScript>{alert('��������20000�����ӣ�');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
hy=rs("��Ա�ȼ�")
if hy=0 then
Response.Write "<script language=JavaScript>{alert('�㲻�ǵȼ���Ա,����й���!');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
else
conn.execute "update �û� set ����=����-20000 where ����='" & aqjh_name & "'" 
conn.execute "update �û� set ����=����+100 where ����='" & to1 & "'"
lenyin="##��%%��ף���߳����ո裬Happy birthday tu you����ף������Ҳ���֣���ף���������֣������н񳯣������н��ա������ţ������ڶ�����һ��ף��%%������֣�%%�ж��ĺʹ��һ��ȹ������õ�һ�죬һ�������ϲѶ����ҹ������ҿ񻶵�4�㻹�������ᣡ" & to1 & "���ĵĳ���<IMG src='img/1-2.gif'>������������<font color=red>100</font>����̾������ϣ����ʱ���á�������"
end if
conn.close
set conn=nothing
end function
%>