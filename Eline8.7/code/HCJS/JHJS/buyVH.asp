<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
wpname=trim(Request.QueryString("wpname"))
if InStr(wpname,"or")<>0 or InStr(wpname,"'")<>0 or InStr(wpname,"`")<>0 or InStr(wpname,"=")<>0 or InStr(wpname,"-")<>0 or InStr(wpname,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT b,h,l,m,i,j FROM b where a='"&wpname&"'",conn
if rs("b")<>"��ͨ" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
tj=rs("l")
yin=rs("h")
jinbi=rs("m")
wpwj=rs("i")
njd=rs("j")
rs.close
rs.open "select ��Ա�ȼ�,����,����ʱ��,��� from �û� where ����='"&sjjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻�߱���������\n["&replace(tj,chr(39),"")&"]��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if

hy=rs("��Ա�ȼ�")
if hy>0 then
	yin=int(yin/2)
end if
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
if yin>rs("����") or jinbi>rs("���") then
	Response.Write "<script Language=Javascript>alert('��ʾ�����򲻳ɹ���ԭ����������������Ҳ���');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "select * from vh where a='"&wpname&"' and b='"&sjjh_name&"'"
if not rs.eof and not rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ������"&wpname&"');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
conn.execute "update �û� set ����=����-" & yin & ",���=���-"&jinbi&",����ʱ��=now() where ����='"&sjjh_name&"'"
conn.execute "insert into vh(a,b,c,d,e,f,g,h,i,l) values('"&wpname&"','"&sjjh_name&"',false,'��',false,'��','"&wpname&"','"&wpwj&"',"&njd&",date())"
rs.close
set rs=nothing
conn.close
set conn=nothing
if hy>0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����Ա������Ʒ������5�ۣ�����"&wpname&"�ɹ��������ڿ���ȥ���ƽ��������ҵĹ����ˣ�');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('��ʾ������"&wpname&"�ɹ�,������ǻ�Ա�������Դ�5�ۣ������ڿ���ȥ���ƽ��������ҵĹ����ˣ�');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
