<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
grade=Int(sjjh_grade)
id=Trim(Request.QueryString("id"))
id2=LCase(id)
if sjjh_name="" then Response.Redirect "../error.asp?id=420"
if grade<7 then Response.Redirect "../error.asp?id=482"
if id2="" then Response.Redirect "../error.asp?id=483"
'total=Application("sjjh_chatrs"&nowinroom)
onlinelist=Application("sjjh_onlinelist"&nowinroom)
ubo=UBound(onlinelist)
dim show()
js=1
for i=1 to ubo step 6
if Instr(LCase(onlinelist(i+1)),id2)<>0 or Instr(onlinelist(i+2),id2)<>0 then
Redim Preserve show(js),show(js+1),show(js+2),show(js+3)
show(js)=onlinelist(i+1)
show(js+1)=onlinelist(i+2)
show(js+2)=onlinelist(i+4)
show(js+3)=onlinelist(i+5)
js=js+4
end if
next
set onlinelist=nothing
totalrec=Int((js-1)/4)
p=1
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m%><html>
<head>
<title>������Ϣ����</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="#000000" face="����" size="6">ը �� �� ��</font></h1>
<a href="javascript:history.go(0)">ˢ��</a> </div>
<hr noshade size="1" color=009900 width="70%">
<div align="center"><b>��������Ա��</b>���� <font color=red><%=total%></font> �ˣ��ҵ����� <font color=red><%=id%></font>
���û��� <font color=red><%=totalrec%></font> ���� </div>
<hr noshade size="1" color=009900 width="70%">
<div align=center>��ҳˢ��ʱ�䣺<font color="#FF0000"><%=sj%></font><br></div>
<table border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="4" bgcolor="E0E0E0" align="center" width="416">
<tr bgcolor="#000000" align="center">
<td><font color="#FFFFFF">���</font></td>
<td><font color="#FFFFFF">�� �� �� ��</font></td>
<td><font color="#FFFFFF">�Ա�</font></td>
<td><font color="#FFFFFF">�� �� ʱ ��</font></td>
<td><font color="#FFFFFF">δ���</font></td>
<td><font color="#FFFFFF">ը��</font></td>
</tr>
<%if grade<3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=4 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=5 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade>6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%>
<tr align="center">
<td><%Response.Write p+j
j=j+1%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=DateDiff("n",show(i+3),sj)%> ��</td>
<td><a href="manbomb.asp?id=<%=server.URLEncode(show(i))%>&ip=<%=show(i+1)%>">��ը</a></td>
</tr>
<%next
end if
end if%>
</table>
<div align="center"><br>
ע�⣺�˹���ֻ������Щ��������ȴ���ڸĵ�����!<br>
�˼�¼�����������ݿ��У��κ�����Ϲը�˵Ķ������ܵ����� </div>
<hr noshade size="1" color=009900 width="70%">
<div align=center class=cp><%Response.Write "���кţ�<font color=blue>" & Application("sjjh_sn") & "</font>����Ȩ����<font color=blue>" & Application("sjjh_user") & "</font><br><font color=999999>" & Application("sjjh_copyright") & "</font>"%></div>
</body>
</html>