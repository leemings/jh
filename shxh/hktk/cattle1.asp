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
	
        if rs("����")>100000 or rs("����")="�ٸ�" then
	response.write "���书�Ѿ���������,��ѡ����������"
	conn.close
	response.end
end if
%>
<%end if%>
<%
	randomize timer
	r=int(rnd*100)
        if r>70 then
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
<form method=POST action='cattle1ok.asp'>
<tr><td align=center>ѡ������������
<select name=nl size=1> 
<option value="0">0������
<option value="10">10������
<option value="20">20������
<option value="30">30������
<option value="40">40������
<option value="50">50������
<option value="60">60������
<option value="70">70������
<option value="80">80������
<option value="90">90������
<option value="100">100������
<option value="110">110������
<option value="120">120������
<option value="130">130������
<option value="140">140������
<option value="150">150������
</select></td></tr>
<tr><td colspan=2 align=center><input type=submit value=ȷ��>
<input type=button value=���� onclick="location.href='xuetang.asp'">
</td></tr>
<tr><td colspan=2 style='font-size:9pt'><hr>
������Ӳ�Զ��ɭ�ֳ�Ⱥ����
ţ�൱���״�ɱ��
�ʺϳ��ڵ�������
��һ�㶯��һ��������=0-10�㡣
</td></tr>
</form>
</table>
</table>

</body>
</html>