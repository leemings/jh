<!--#include file="jha.asp"-->
<%
name=Session("Ba_jxqy_username")
sql="select * from �û� where ����='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
  MsgBox "�㲻�ǽ�������"
  location.href = "javascript:history.back()"
</script>
<%
else
	
        if rs("����")>10000 or rs("����")="�ٸ�" then
	response.write "���书�Ѿ���������,��ѡ����������"
	conn.close
	response.end
end if
%>
<%end if%>
<%
	randomize timer
	r=int(rnd*100)
        if r>80 then
%>
<script language=vbscript>
  MsgBox "û���ҵ�����"
  location.href = "javascript:history.back()"
</script>
<%end if%>
<html>
<TITLE>����ѵ����</TITLE>
<head>
<style>
input, body, select, td{font-size:14;line-height:160%}
</style>
</head>

<body oncontextmenu=self.event.returnValue=false>
<p align=center style='font-size:16;color:yellow'>��ӭ<%=session("myname")%>����ѵ��
<table border=1 bgcolor="#FFCC99" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#BEE0FC>
<table>
<tr><td>
<form method=POST action='cat1ok.asp'>
<tr><td align=center>ѡ������������
<select name=nl size=1> 
<option value="0">0������
<option value="10">100������
<option value="20">200������
<option value="30">300������
<option value="40">400������
<option value="50">500������
<option value="60">600������
<option value="70">700������
<option value="80">800������
<option value="90">900������
<option value="100">1000������
</select></td></tr>
<tr><td colspan=2 align=center><input type=submit value=ȷ��>
<input type=button value=���� onclick="location.href='xuetang.asp'">
</td></tr>
<tr><td colspan=2 style='font-size:9pt'><hr>
�����è���ٶȷǳ��죬
�����С������èצ��
��è�Ĺ���������, ���ʺ���������������=0-5�㡣
</td></tr>
</form>
</table>
</table>

</body>
</html>