<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�����ϳ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
says="<font color=red>�������ϳɡ�<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�����ϳ�
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,w4,ְҵ,��� FROM �û� WHERE ����='"&aqjh_name&"'",conn
dj=rs("�ȼ�")
fla=rs("����")
if rs("����")<75000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��75000�㰡��');window.close();}</script>"
	response.end
end if
if rs("ְҵ")<>"����ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊ����ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("���")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ľ�Ҳ����޷�ʩչѽ������Ҳ��2������');window.close();}</script>"
	response.end
end if
if dj<60 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ[60]��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w4")
if isnull(duyao) or duyao=abate(rs("w4"),"÷����",5) or duyao=abate(rs("w4"),"������",5) or duyao=abate(rs("w4"),"��Ѫ����",5) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����������ȱһ���ɿ��������!');</script>"
	response.end
end if

rs.close
rs.open "select w4 from �û� where ����='"&aqjh_name&"'",conn
duyao=abate(rs("w4"),"÷����",5)
conn.execute "update �û� set  w4='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=abate(rs("w4"),"������",5)
conn.execute "update �û� set  w4='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=abate(rs("w4"),"��Ѫ����",5)
conn.execute "update �û� set  w4='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=add(rs("w4"),"�����绨��",2)
conn.execute "update �û� set ���=���-2,w4='"&duyao&"' where ����='"&aqjh_name&"'"
peibashi=aqjh_name & "<bgsound src=wav/m2.mid loop=1>����ʦȡ��÷���ڡ������ڡ���Ѫ�������ֳ�ǿ���������ְ���������һ��һ����ɫ��â�������ְ������һ��������һ��������ʧ����Ӱ���٣�ֻ�����������˵���ǧ��֮�������İ���֮��<font color=red>�����绨��</font>2ֻ," & aqjh_name & "��Ե�ɺ�֮�»�õ������п���ʹ��������������Ǻ���������....."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>%>