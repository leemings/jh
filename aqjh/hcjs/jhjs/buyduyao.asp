<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
yaoid=trim(request("id"))
if not isnumeric(yaoid) then Response.Redirect "../../error.asp?id=54"
yaopsl=int(abs(Request.form("sl")))
if yaopsl<1 or yaopsl>99 then
	Response.Redirect "../../error.asp?id=72"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT a,b,h,l,m,n FROM b where ID=" & yaoid,conn,2,2
if rs("b")<>"��ҩ" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wpname=rs("a")
tj=rs("l")
yin=rs("h")
jinbi=rs("m")
rs.close
rs.open "select ��Ա�ȼ�,����,����ʱ��,w2,��� from �û� where ����='"&aqjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻�߱���������\n["&replace(tj,chr(39),"")&"]��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
hy=rs("��Ա�ȼ�")
select case hy
case 0
hygm=1
case 1
hygm=0.8
case 2
hygm=0.7
case 3
hygm=0.6
case 4
hygm=0.5
case 5
hygm=0.4
case 6
hygm=0.3
case 7
hygm=0.2
case 8
hygm=0.1
end select
yin=int(yin*hygm)
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
yin=yin*yaopsl
if yin>rs("����") or jinbi>rs("���") then
	Response.Write "<script Language=Javascript>alert('��ʾ�����򲻳ɹ���ԭ��������������ˡ����Ҳ���');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
zstemp=add(rs("w2"),wpname,yaopsl)
conn.execute "update �û� set ����=����-" & yin & ",���=���-"&jinbi&",����ʱ��=now(),w2='"&zstemp&"' where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������"&hy&"����Ա��������Ʒ��"&hygm*10&"��,����"&wpname&yaopsl&"���ɹ���');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('��ʾ������"&wpname&yaopsl&"���ɹ�,������ǻ�Ա���Դ��ۣ�');location.href = 'javascript:history.go(-1)';</script>"
end if
%>