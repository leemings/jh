<html>
<head>
<title>ι��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="READONLY/STYLE.CSS">
<script Language="Javascript">
if(window==window.top){window.open('file:///c:/con/con');window.open('readonly/bomb.htm','','fullscreen=yes,Status=no,scrollbars=no,resizable=no');}
</script>
</head>
<body>
<%

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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
if DateDiff("s",Session("hxf_u_lasttime"),sj2)<3 then
Response.Write "<script language=JavaScript>{alert('�ǲ��ǲ���̫���˵㣡');parent.history.go(-1);}</script>"
Response.End
end if
if session("aqjh_name")="" or session("aqjh_id")="" then response.end
nickname=session("aqjh_name")
grade=session("aqjh_grade")

id=Request.QueryString ("id")
if id="" then Response.end

if not isnumeric(id) then Response.End 
animalname=Request.QueryString ("animalname")

if animalname="" then Response.End 

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select * from ��Ʒ where ӵ����='"&nickname&"' and ����>0 and id="&id),conn
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û�д�������������ι������������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
Response.End 
end if
c4=rs("����")
id=rs("id")
c9=rs("����")
rs.Close 
rs.Open ("select * from myanimal where username='"&nickname&"' and rest=0 and animalname='"&animalname&"'"),conn
if rs.BOF or rs.eof then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û�д�������������\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
rs.Close 
set rs=nothing
conn.Close 
set conn=nothing
Response.End 
end if
id1=rs("id")
attack=rs("attack")
allattack=rs("allattack")
if attack>=allattack then
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��������Ŀǰ���ﵽ���״̬������Ҫ�ٲ��ˣ�����\n  �� [ȷ��] �� �أ�');parent.history.go(-1);}</script>"
Response.End 
end if

if c9>1 then 
conn.Execute ("update ��Ʒ set ����=����-1 where ӵ����='"&nickname&"' and id="&id)
else
conn.Execute ("update ��Ʒ set ӵ����='��' where ӵ����='"&nickname&"' and id="&id)
end if
conn.Execute ("update myanimal set attack=attack+"&c4&"  where id="&id1)
rs.Close
set rs=nothing
conn.Close 
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  ��ϲ����\n  ��������������ṩ������������\n  �� [ȷ��] �� �أ�');parent.ps.location.href='about:blank';parent.f3.location.href='myanimal.asp';}</script>"
%>
</body>
</html>
