<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ħ��ʯ
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
says="<font color=red>��ħ��ʯ��<font color=" & saycolor & ">"+lbsw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function lbsw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,�ȼ�,allvalue,ְҵ,����,����,ת�� FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<1 or fn1>10 then
Response.Write "<script language=JavaScript>{alert('ħ��ʯһ������1��಻�ܳ���10����');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ʹ��ħ��ʯ������ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
duyao=rs("w6")
if isnull(duyao) or duyao=abate(rs("w6"),"ħ��ʯ",fn1*10)  then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������"&fn1*10&"��ħ��ʯҲû��Ү.���������!');</script>"
	response.end
end if
if rs("�ȼ�")<30 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ñ����㻹��������Ҫ��30�����ϣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("ת��")>2 then
	Response.Write "<script language=JavaScript>{alert('���Ѿ��Ƕ��ת�����ˣ�������������ȥ�չ����˰ɣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>10  then
	Response.Write "<script language=JavaScript>{alert('��̫̰�ˣ����𳬹�10����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<fn1*150000 then
	Response.Write "<script language=JavaScript>{alert('�㷨������"&fn1*150000&"��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<fn1*10000 then
	Response.Write "<script language=JavaScript>{alert('����²���"&fn1*10000&"��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select allvalue,����,����,�ȼ�,w6 FROM �û� WHERE ����='" & aqjh_name &"'" ,conn
if rs("�ȼ�")>200 then
lbsw=aqjh_name & " ����ܻ��ֹ��ߵ��ˣ���Ҫ���𣿺ٺٺ�.(����˵ȼ�����200)��"
else
duyao=abate(rs("w6"),"ħ��ʯ",fn1*10)
conn.execute "update �û� set w6='"&duyao&"',allvalue=allvalue+"&fn1*1000&",����=����-"&fn1*150000&",����=����-"&fn1*10000&" where ����='" & aqjh_name &"'"
lbsw="����<font color=red size=2>" & aqjh_name & "</font>ʹ���˴�˵�е���������֮����<font color=blue>ħ��ʯ</font>���ܻ��ֿ���<font color=red>"&fn1*1000&"</font>�㣬�����½�<font color=red>"&fn1*10000&"</font>��,��������<font color=red>"&fn1*150000&"</font>�㣬���������ڶ�С������ֱ�������������Ŀֲ�ħ��������"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>%>