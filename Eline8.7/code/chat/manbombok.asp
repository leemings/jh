<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc/func.asp"-->
<%Response.Expires=0
inthechat=Session("sjjh_inthechat")
userip=Request.ServerVariables("REMOTE_ADDR")

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
grade=Int(sjjh_grade)
if grade<7 then Response.Redirect "../error.asp?id=482"
'У���ʱ�Ƿ񱻽���
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="SELECT grade FROM �û� WHERE  ����='"&sjjh_name&"'"
rs.open sql,conn,1,3
if Not(rs.Eof and rs.Bof) then
grade=rs("grade")
if intherchat<>"1" and grade<7 then Response.Redirect "../error.asp?id=482"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
'���У��
if inthechat<>"1" and sjjh_name<>"СС��" then Response.Redirect "manerr.asp?id=482"
bombname=Server.HTMLEncode(Trim(Request.Form("bombname")))
bombwhy=Server.HTMLEncode(Trim(Request.Form("bombwhy")))
logok=Trim(Request.Form("logok"))
if bombname="" then Response.Redirect "manerr.asp?id=222"
if bombwhy="" then Response.Redirect "manerr.asp?id=224"
if CStr(bombname)=CStr(sjjh_name) then Response.Redirect "manerr.asp?id=223"
if len(bombwhy)>60 then bombwhy=Left(bombwhy,60)
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
onlinelist=Application("sjjh_onlinelist"&nowinroom)
cz=0
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if CStr(onlinelist(i+1))=CStr(bombname) then
cz=1
bombip=onlinelist(i+2)
end if
next
if cz=0 then Response.Redirect "manerr.asp?id=225&bombname=" & server.URLEncode(bombname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="SELECT grade FROM �û� WHERE ����='" & bombname & "'"
rs.open sql,conn,1,3
if sjjh_grade<=rs("grade") then
	rs.close
	set rs=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������ԶԸ߼�����Ա������');window.close();</script>"
	response.end
end if
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& bombname &"','����','"&bombwhy&"')"
rs.close
set rs=nothing
bomblog="�� <font color=red>" & bombname & "</font>(<font color=blue>" & bombip & "</font>) ը�ɣ�<br><font color=009900>��ԭ��" & bombwhy & "��</font>"
if logok="1" then
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
end if
conn.close
set conn=nothing
Session("sjjh_lasttime")=sj
Application.Lock
sd=Application("sjjh_sd")
line=int(Application("sjjh_line"))
Application("sjjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
	sd(192)="660099"  	'ɫ��
	sd(193)="��Ϣ" 		'����
	sd(194)=sjjh_name	'˵����
	sd(195)=""			'����
	sd(196)="���"		'�ܻ���
	sd(197)="<font color=black>����ը��</font><font color=8800FF><font color=red>" & sjjh_name & "</font>���������ذ�<font color=red>" & bombname & "</font>ը�����á�����ԭ��" & bombwhy & "��</font><font class=t>(" & t & ")</font>"		'��������
	sd(198)=sjjh_id		'�Լ�id
	sd(199)=0			'�Է�id
	sd(200)=nowinroom	'���ڷ���
	Application("sjjh_sd")=sd
	Application("sjjh_bombname")=Application("sjjh_bombname") & " " & bombname & " "
Application.UnLock
call boot(bombname,"��ը�������ߣ�"&sjjh_name&","&bombwhy)
%><html>
<head>
<title>ը������</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF">��ը��������</font></h1>
</div>
<hr noshade size="1" color=009900>
<b>��������ɡ�</b> <br>
ը���Ѿ��ӳ���������ͻῴ���Է����ص�ˤ����ȥ����ղŵĲ���<%if logok="0" then Response.Write "<font color=red>û��</font>"%>����¼�ڹ������С�
<hr noshade size="1" color=009900>
<br>
<table border="1" cellspacing="0" cellpadding="10" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td>
<p><%=sj%></p>
<p><%=sjjh_name&"("&userip&")"%></p>
<p><%=bomblog%></p>
<div align="center">
<input type="button" value="����" onclick="javascript:history.go(-2)">
</div>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900>
<div align=center class=cp><%Response.Write "���кţ�<font color=blue>" & Application("sjjh_sn") & "</font>����Ȩ����<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>
<html></html>