<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select  * from �û� where ����='"&aqjh_name&"'",conn
allvalue=rs("allvalue")
bigdk=allvalue*100
yinliang=rs("����")
jhdenji=rs("�ȼ�")
nowyl=rs("����")
nowck=rs("���")
rs.close
rs.open "select * from k where d=false and DateDiff('d',b,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	do while not (rs.bof or rs.eof)
		name=rs("a")
		conn.Execute ("update k set d=true where a='"&name&"'")
		conn.Execute ("delete * from �û� where ����='"& name &"'")
		conn.Execute ("insert into l(b,a,c,d,e) values ('" & name & "',"&sqlstr(sj)&",'������ɱ��','Ƿծ��Ǯ��ûǮҪ����','����')")
		rs.movenext
		name=""
		aqjh_name=""
	loop
end if
%>
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" background="../../bg.gif">
<p align="center"><font size="6" face="����">������������</font></p>
<p align="center">&nbsp;</p>
<form method="post" action="dk.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#E0E0E0">
<tr>
<td>�����:</td>
<td><%=aqjh_name%></td>
</tr>
<tr>
<td>����������</td>
<td><%=nowyl%>��</td>
</tr>
<tr>
<td>���ڴ��:</td>
<td><%=nowck%>��</td>
</tr>
<tr>
<td>��������:</td>
<td><%=bigdk%>��</td>
</tr>
<tr>
<td>�����</td>
<%
rs.close
rs.open "select * from k where a='"&aqjh_name&"' and d=false",conn
if rs.BOF or rs.EOF then
	%>
	<td>
	<%if jhdenji>3 then%>
		<input type="text" name="dk" size="10" maxlength="10">
	<%else%>
		<font color=red>���ܲ���</font>
	<%end if%>
	</td>
	</tr>
<tr>
<td colspan="2">
<div align="center">
<%if jhdenji>3 then%>
	<input type=submit value="����" name="submit">
	<input type="reset" name="reset" value="���">
<%else%>
	<font color=red>ս��С��3��������������㣡</font>
<%end if%>
</div>
</td>
<%else%>
<td>
<%if yinliang>rs("c")*1.5 then%>
<input type="text" name="dk" size="10" maxlength="10" value='<%=rs("c")%>' readonly>
<%else%>
<font color=red>���ܲ���</font>
<%end if%>
<br>�������ڣ�<%=rs("b")%>
</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<%if yinliang>rs("c")*1.5 then%>
<input type=submit value="����" name="submit">
<input type="reset" name="reset" value="���">
<%else%>
<font color=red>����ֽ𲻹�����!</font>
<%end if%>
</div>
</td>
<%end if
rs.close
set rs=nothing
conn.Close
set conn=nothing%>
</tr>
</table>
</form>
<p align="center">��Ǯׯ�ṩ���С���ﲻ����ս���ȼ�����3�����Ŵ���<br>
����������һ������7�졣�׻�˵&quot;Ƿծ��Ǯ��ûǮ����&quot;�����ڲ����߱�ׯ����<br>
ɱ�˽���ɱ֮��<font color=red>(��ɾ������ID)</font>����λ�����໥ת�棡���������� <br>
(��������Ǯ����Ϊ��100������150�����������ĺ�ѽ������׬Ǯ��ѽ��) </p>
<p align="center"><font color=red>�ڰ��齭����վ��</font></p>
</body>
</html>
<%
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
%>