<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
pai=lcase(request("pai"))
a=lcase(request("a"))
wgname=trim(request.form("wgname"))
wg=int(abs(request.form("wg")))
nl=int(abs(request.form("nl")))
if instr(wgname,"'")<>0 then
	Response.Write "<script Language=Javascript>alert('���ѽ���ڿ���������ز����˰ɣ���');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if (wg+nl)>10000 then
	Response.Write "<script Language=Javascript>alert('��ʾ�������书��������֮�ϲ����Դ���1��');location.href = 'javascript:history.back()';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM p where b='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ѽ���ڿ���������ز����˰ɣ���');location.href = 'javascript:history.back()';</script>"
	response.end
end if
rs.close
if a="m" then
	id=int(abs(request.form("id")))
	conn.Execute "update y set a='" & wgname & "',c="& wg &",d=" & nl & " where b='"&pai&"' and id="&id
end if
if a="n" then
	conn.Execute "insert into y(a,b,c,d,e) values ('"&wgname&"','"&pai&"',"&wg&","&nl&",0)"
end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "setwg.asp"
%>