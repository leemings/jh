<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if 
id=request("id")
if not isnumeric(id) then Response.Redirect "../../error.asp?id=54"
'sl=int(abs(Request.form("sl")))
sl=1
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where ID=" & id,conn
if rs("b")<>"����" and rs("b")<>"����" and rs("b")<>"����" and rs("b")<>"ͷ��" and rs("b")<>"˫��" and rs("b")<>"װ��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wu=rs("a")
yin=rs("h")
lx=rs("b")
gj=rs("f")
fy=rs("g")
sm=rs("b")
rs.close
rs.open "select ��Ա�ȼ�,����,����ʱ�� from �û� where id="&aqjh_id,conn
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
if yin*sl>rs("����") then
	Response.Write "<script Language=Javascript>alert('��ʾ�����򲻳ɹ���ԭ���������������');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update �û� set ����=����-" & yin*sl & ",����ʱ��=now() where id="&aqjh_id
rs.close
'��ʼ����
rs.open "select id from w where a='"& wu&"' and b="&aqjh_id
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into w (a,b,c,g,h,i,l,d) values ('"&wu&"',"& aqjh_id&",'"& lx&"',"& gj &","& fy &","&sl&","&int(yin)&",'"&sm&"')"
else
	id=rs("id")
	conn.execute "update w set l=" & int(yin) & ",i=i+" & sl & ",d='"&sm&"' where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('��ʾ������"&hy&"����Ա��������Ʒ��"&hygm*10&"��,����"&wu&sl&"���ɹ���');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('��ʾ������"&wu&sl&"���ɹ�,������ǻ�Ա���Դ��ۣ�');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
