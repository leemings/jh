<%
username=session("yx8_mhjh_username")
usergrade=session("yx8_mhjh_usergrade")
usercorp=session("yx8_mhjh_usercorp")
if username="" then Response.Redirect "../../error.asp?id=016"
if usercorp<>"�ٸ�" then Response.Redirect "../../exit.asp"
dim conn,param,action
Set conn = Server.CreateObject("ADODB.Connection")
param = "driver={Microsoft Access Driver (*.mdb)};"
conn.Open param & ";dbq=" & Server.MapPath("qwas.asp")
action=request.querystring("action")


if action="add" then
dim lx,question,A,B,C,D,answer
lx=request.form("lx")
question=request.form("question")
A=request.form("A")
B=request.form("B")
C=request.form("C")
D=request.form("D")
answer=request.form("ANSWER")
question=replace(question,"'","''")
a=replace(a,"'","''")
b=replace(b,"'","''")
c=replace(c,"'","''")
d=replace(d,"'","''")

sql="Select * from QUESTION Where QUESTION='"&QUESTION&"'"
Set rs = conn.Execute( sql )
If RS. EOF AND rs.BOF Then
str="Insert Into QUESTION (LX,QUESTION,A,B,C,D,ANSWER) Values('"&LX&"','"&QUESTION&"','"&A&"','"&B&"','"&C&"','"&D&"','"&ANSWER&"')"
conn.execute(str)
set rs=server.createobject("adodb.recordset")
set rs=conn.execute("Select * from QUESTION Where QUESTION='"&QUESTION&"'")
%>
<script language="javaScript">
alert ("<%=question%>\n\n����ɹ���")
location.href="question.asp"
</script>
<%end if %>
<script language="javaScript">
alert ("<%=question%>\n\n�������ݿ�������ͬ�����⡣")
location.href = "question.asp"
</script>
<%end if
if action="edit" then
sql="Select * from QUESTION where id="&request.querystring("id")
Set rs = conn.Execute( sql )
lx=rs("lx")
question=rs("question")
A=rs("A")
B=rs("B")
C=rs("C")
D=rs("D")
answer=rs("ANSWER")

question=replace(question,"'","''")
a=replace(a,"'","''")
b=replace(b,"'","''")
c=replace(c,"'","''")
d=replace(d,"'","''")
%>

<html>
<head>
<title>�޸�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="home.css">

</head>
<body background="../../images/newbg.gif"  >
<p>�޸�����</p>
<form name="form" action="action.asp?action=updata" method="post" >
<table width="32%" border="0" cellspacing="0" cellpadding="2">
<tr>
<td width="9%" nowrap>��������</td>
<td width="91%">
<input type="radio" name="lx" value="1" checked>
1����
<input type="radio" name="lx" value="2">
2����
<input type="radio" name="lx" value="3">
3��ѧ
<input type="radio" name="lx" value="4">
4��ѧ
<input type="radio" name="lx" value="5">
5�ۺ� </td>
</tr>
<tr>
<td width="9%" nowrap>��Ŀ</td>
<td width="91%">
<input type="text" name="question" size="50" maxlength="100" value="<%=question%>">
</td>
</tr>
<tr>
<td width="9%" nowrap>��A</td>
<td width="91%">
<input type="text" size="50" maxlength="100" name="A" value="<%=a%>">
</td>
</tr>
<tr>
<td width="9%" nowrap>��B</td>
<td width="91%">
<input type="text" name="B" size="50" maxlength="100" value="<%=b%>">
</td>
</tr>
<tr>
<td width="9%" nowrap>��C</td>
<td width="91%">
<input type="text" name="C" size="50" maxlength="100" value="<%=c%>">
</td>
</tr>
<tr>
<td width="9%" nowrap>��D</td>
<td width="91%">
<input type="text" name="D" size="50" maxlength="100" value="<%=d%>">
</td>
</tr>
<tr>
<td width="9%" nowrap>��ȷ��</td>
<td width="91%">
<input type="radio" name="ANSWER" value="A" checked>
A
<input type="radio" name="ANSWER" value="B">
B
<input type="radio" name="ANSWER" value="C">
C
<input type="radio" name="ANSWER" value="D">
D </td>
</tr>
<tr>
<td width="9%" nowrap>&nbsp;</td>
<td width="91%">
<input type="submit" name="Submit" value="�޸�">
</td>
</tr>
</table>
</form>
</body>
</html>
<%end if %>
<%
lx=request.form("lx")
question=request.form("question")
A=request.form("A")
B=request.form("B")
C=request.form("C")
D=request.form("D")
answer=request.form("ANSWER")
question=replace(question,"'","''")
a=replace(a,"'","''")
b=replace(b,"'","''")
c=replace(c,"'","''")
d=replace(d,"'","''")


if action=updata then
sql="updata  QUESTION set id="&lx&",question="&question&",a="&A&",B="&B&",C="&C&",D="&D&",ANSWER="&ANSWER&""
SET RS=CONN.EXECUTE(SQL)%>
<script language="javaScript">
alert ("<%=question%>\n\n�޸ĳɹ�����")
location.href = "javascript:window.close()"
</script>

<%
END IF
%>

