<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
who=Trim(Request.Form("who"))
if who="" then who=aqjh_name
show=Split(Trim(useronlinename),"  ",-1)
x=UBound(show)
chatroombgimage=Application("aqjh_chatimage")
chatroombgcolor=Session("afa_chatbgcolor")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ���� FROM �û� where ����='" & aqjh_name &"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�ǽ������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
pai=rs("����")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>������</title>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'>����������</span></font></font><br>
</font> </p>
<p class=p1 align="center"><font color="#FFFFFF">����������<br>
�˺�=�����˺��ܺ�</font></p>
<form method=POST action='stunt1.asp'><font color="#FF0033">��������</font><br>
<select name="to1">
<%
for i=0 to x
	if show(i)<>aqjh_name and show(i)<>peiou then
		%><option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option><%
	end if
next
%>
</select>
<font color="#FFFFFF"><br>
<br>
����1<br>
<select class="smallSel" name="at1" size="1">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"' >"+rs("a")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
����2<br>
<select class="smallSel" name="at2">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"'>"+rs("a")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
����3<br>
<select class="smallSel" name="at3">
<%
rs.close
rs.open "SELECT * FROM y where b='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("a"))+"'>"+rs("a")+""+chr(13)+chr(10)
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</select>
<br>
<br>
</font><font color=#ffffff>
<input type=submit value="�� ��" name=submit>
</font><br>
</form>
</div>
</body></html>