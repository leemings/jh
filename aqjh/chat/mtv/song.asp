<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/czdj.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if aqjh_name="" or Session("aqjh_inthechat")<>"1" or Instr(useronlinename," "&aqjh_name&" ")=0 then Response.Redirect "../chaterr.asp?id=001"
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
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set Rs=conn.Execute("select  * FROM song")
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>���</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<script language="javaScript">
function send(){sendurl="songplay.asp?name=" + document.forms[0].song.options[document.forms[0].song.selectedIndex].value;window.location.href=sendurl;}
</script>
<link rel="stylesheet" href="../readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<div align=center><font color="#FFFF00" style="font-size:9pt">���ϵͳ | <a href=swflist.asp><font color="#FFFF00" style="font-size:9pt">MTV�㲥</font></a></font></div>
<hr size=1 color=FFFF00>
<table border="0" width="100%">
<form method="post" action="songplay.asp?" name="" target="ps">
<tr>
<td>
<p> <font color="#FFFF00" style="font-size:9pt">ѡ��MP3�б�:</font><br>
<select name="song" style="font-size:9pt" size="1">
<%do while not rs.bof and not rs.eof %>    
<option value="<%=rs("name")%>"><%=rs("name")%></option>
<%
rs.movenext
loop
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
erase mid    
</select>
</p>
<p><font color="#FFFF00" style="font-size:9pt">���ŷ�ʽ��</font><br>
<input type="radio" name="loopok" value="1" checked>
<font style="font-size:9pt">ֻ��һ��<br>    
<input type="radio" name="loopok" value="infinite">
<font style="font-size:9pt">��������<br>    
<input type="radio" name="loopok" value="ddj">
<font style="font-size:9pt">������<br>    
&nbsp;&nbsp;&nbsp; ��Ҫ<%=xyyl%>�����<br>    
<input type="radio" name="loopok" value="dhy">
<font style="font-size:9pt">�������<br>    
&nbsp;&nbsp;&nbsp; ��Ҫ<%=pyyl%>�����<br>    
<br>
ף������<br>
<input type="text" name="zhufu" size="17" value="ף���ÿ�춼��һ��������!" maxlength="50"><br>
��������<br>
    <select name="to1">
      <%useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
for i=0 to x
if show(i)<>aqjh_name Then
%>
      <option value="<%=show(i)%>" selected><%=show(i)%></option>
<%end if
next%>
<%if x=0 then%>
      <option value="�����Լ�һ����" selected>�����Լ�һ����</option>
<%end if%>
    </select>
 ��</p>
<table border="0" cellpadding="4">
<tr>
<td>
<input type="submit" name="play" value="����">
</td>
<td align="right">
<input type="submit" name="st" value="ֹͣ">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body>
</html>