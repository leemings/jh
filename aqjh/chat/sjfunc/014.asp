<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'��ħ��ʯ
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
says="<font color=red>����ħ��ʯ��<font color=" & saycolor & ">"+peibashi(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��ħ��ʯ
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,w6,ְҵ,��� FROM �û� WHERE ����='"&aqjh_name&"'",conn
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
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ2�����ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<30 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ[30]��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"˧����",2) or duyao=abate(rs("w6"),"������",2) or duyao=abate(rs("w6"),"���黷",2) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����������ȱһ���ɿ��������!');</script>"
	response.end
end if

rs.close
rs.open "select w6 from �û� where ����='"&aqjh_name&"'",conn
duyao=abate(rs("w6"),"˧����",2)
conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=abate(rs("w6"),"������",2)
conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=abate(rs("w6"),"���黷",2)
conn.execute "update �û� set  w6='"&duyao&"' where ����='"&aqjh_name&"'"
duyao=add(rs("w6"),"ħ��ʯ",1)
conn.execute "update �û� set ���=���-2,w6='"&duyao&"' where ����='"&aqjh_name&"'"
peibashi=aqjh_name & "<bgsound src=wav/m2.mid loop=1>ȡ��˧�����������黷���ֱ�����ֱ�������һ��һ����â<img src=img/hs1.gif>�������ֱ��ｫ��ع����ظ�������" & aqjh_name & "��Ե�ɺ�֮�»��������ħ���������µ�-<font color=red>ħ��ʯ</font>-�����ǲȵ���ʺ����....."
 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>%>