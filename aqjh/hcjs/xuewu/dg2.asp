<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
my=aqjh_name
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ����ܽ��뽭�����"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"������ѧ��ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
tl=rs("����")
nl=rs("����")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
if tl<300 or nl<350 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "������Ŀǰ������������ֵ��������ȥ��������������������ɣ�"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update �û� set ����=����-300,����=����-350,�书=�书+60,����=����+10,����=����+2,����ʱ��=now() where ����='"&aqjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
message="" & my & "�������ѧ��һ�죬���书+60����+10����+2����������300����350��"
says="<font color=red>��������ݡ�</font><font color=green>["&aqjh_name&"]�������ѧ��һ��[����ʮ��ץ]���书+60����+10,����+2,��������300����350</font>"
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
end if
end if%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>�������:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="�� ��" onclick="location.href='xuetang.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>
