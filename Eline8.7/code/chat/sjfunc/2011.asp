<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'�������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>�������<font color=" & saycolor & ">"+mnl(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function mnl(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select �Ա�,w6,�ȼ�,����,ְҵ,���� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn,2,2
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ʹ��ҡǮ��������ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"������",1)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����û��������Ү.���������!');</script>"
	response.end
end if
if rs("�ȼ�")<80 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ñ����㻹��������Ҫ��80�����ϣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<3000 then
	Response.Write "<script language=JavaScript>{alert('�㷨������3000���ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w6,����,�Ա� FROM �û� WHERE ����='" & sjjh_name &"'" ,conn
if rs("�Ա�")="Ů" then
mnl=sjjh_name & " ����һЦ��������Ů�������˺��ô�.�ó��������ˣ�"
else
duyao=abate(rs("w6"),"������",1)
conn.execute "update �û� set �Ա�='Ů',w6='"&duyao&"',����=����-3000 where ����='" & sjjh_name &"'"
mnl="<font color=red size=2>" & sjjh_name & " ͻȻ������ <font color=blue>������</font>!" & sjjh_name & " �ڲ�֪�����б��Ư����Ů����....</font>"

end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>