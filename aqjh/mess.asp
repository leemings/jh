<%@ LANGUAGE=VBScript codepage ="936" %>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")%>
<html>
<head>
<title>������Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#88AFD7">
<%
if aqjh_name<>"" then
rs.open "select ʦ��,ʦ����Ǯ,���� from �û� where ����='" & aqjh_name & "'",conn
sf=rs("ʦ��")
Response.Write "<div align='center'><font size=-1>ʦ����"& sf &"</font></div>"
if rs("ʦ��")<>"" and rs("ʦ��")<>"��" and rs("ʦ����Ǯ")<>"��Ǯ"&Day(date()) then
	yin=int(rs("����")*0.05)
	conn.execute "update �û� set ����=����-"& yin &",ʦ����Ǯ='��Ǯ"& Day(date()) &"' where ����='"& aqjh_name &"'"
	conn.execute "update �û� set ����=����+"& yin &" where ����='"& sf &"'"
	Response.Write "<div align='center'><font size=-1>"& aqjh_name &"�����Ͻ�ʦ��"& yin &"����������Т��ʦ����"& sf &"</font></div>"
end if
rs.close
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
rs.open "select b,a from k where d=false and a='"& aqjh_name & "'",conn
if not(rs.BOF or rs.EOF) then
	Response.Write "<br>���ڽ���Ǯװ�д���������ڣ�"& rs("b") &"ע�⻹��,7�첻��ɾ��ID��"
end if
rs.close
rs.open "select * from k where d=false and a='"& aqjh_name & "' and DateDiff('d',b,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	name=rs("a")
	conn.Execute ("update k set d=true where a='"&name&"'")
	conn.Execute ("delete * from �û� where ����='"&name&"'")
	'conn.Execute ("insert into l(b,a,c,e,d) values ('" & name & "',"&sqlstr(sj)&",'������ɱ��','Ƿծ��Ǯ��ûǮҪ����','����')")
	Response.Write "��Ϊ��7��û�л��������ID��ɾ���ˣ�"
	aqjh_name=""
	Session.Abandon
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
else%>
<font size=-1><%=Application("aqjh_chatroomname")%>��ӭ����</font> 
<%end if
%>
<div align="center"></div>
</body></html>