<%@ LANGUAGE=VBScript codepage ="936" %>
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
vhid=trim(Request.Form("id"))
vhname=trim(Request.Form("vhname"))	 'g
jrtg=trim(Request.Form("jrtg"))		 'd or f
jr=Request.Form("jr")					 'c or e
jrtc=Request.Form("jrtc")
vhname=replace(vhname," ","")
jrtg=replace(jrtg," ","")

if not IsNumeric(vhid) or InStr(vhname,"or")<>0 or InStr(vhname,"'")<>0 or InStr(vhname,"`")<>0 or InStr(vhname,"=")<>0 or InStr(vhname,"-")<>0 or InStr(vhname,",")<>0 or InStr(vhname,"<")<>0 or InStr(vhname,">")<>0 or InStr(vhname,"[")<>0 or InStr(vhname,"]")<>0 or InStr(jrtg,"or")<>0 or InStr(jrtg,"'")<>0 or InStr(jrtg,"`")<>0 or InStr(jrtg,"=")<>0 or InStr(jrtg,"-")<>0 or InStr(jrtg,",")<>0 or InStr(jrtg,"<")<>0 or InStr(jrtg,">")<>0 or InStr(jrtg,"[")<>0  or InStr(jrtg,"]")<>0  or InStr(jr,"or")<>0 or InStr(jr,"'")<>0 or InStr(jr,"`")<>0 or InStr(jr,"=")<>0 or InStr(jr,"-")<>0 or InStr(jr,",")<>0 then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if InStr(jrtg,"##")=0 or InStr(jrtg,"$$")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����빫���б��������##���͡�$$����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT b FROM vh where id="&vhid,conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ʲô���뵷���𣿣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("b")<>aqjh_name then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô������ֲ�����ģ��뵷���𣿣�');}</script>"
	response.end
end if
rs.close
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<5 then
	s=5-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
conn.execute "update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
if jrtc then
	conn.execute "update vh set c=false where b='"&aqjh_name&"'"
	conn.execute "update vh set c="&jr&",d='"&jrtg&"',g='"&vhname&"' where id="&vhid
else
	conn.execute "update vh set e=false where b='"&aqjh_name&"'"
	conn.execute "update vh set e="&jr&",f='"&jrtg&"',g='"&vhname&"' where id="&vhid
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if jrtc then
Response.Write "<script Language=Javascript>alert('��ʾ�����½��뽭�������ҹ���ɹ���');location.href = 'javascript:history.go(-1)';</script>"
else
Response.Write "<script Language=Javascript>alert('��ʾ�������˳����������ҹ���ɹ���');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
