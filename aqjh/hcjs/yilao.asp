<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<html>
<head>
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%>������</title></head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='"&aqjh_name&"'" ,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "�Բ��𣬲��ܽ��в������㻹���ǽ������ˣ�����"
window.close()
</script>
<%
else
sm=rs("����")
conn.close
set conn=nothing
set rs=nothing

%>
<font color="#000000">��ӭ<%=aqjh_name%>�����������ƣ�ҽ���¹ʲ������Σ� </font> </div>
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table>
<form method=POST action='yilao2.asp' onsubmit="this.ok.disabled=true">
<tr>
<td>��ҽ��˵�����������Ϊ<%=sm%>��
<%
if sm>=1000 then response.write "����Ҫ���ƣ�":p=1
if sm<1000 and sm>0 then response.write "������Щ��Ʒ��":p=2
if sm<0 and sm>-500 then response.write "�����������������ƣ�":p=3
if sm<-500 then response.write "���˵ò��ᣬ�粻���ƽ�Σ��������":p=4
%>
<tr>
<td>���Ʒ�����<input name="yilao" readonly size="8"
value="<%
if p=1 then response.write "��"
if p=2 then response.write "��Щ��Ʒ"
if p=3 then response.write "һ������"
if p=4 then response.write "��������"
%>"> <input name="ok" type="submit" value="ȷ��"> <input type="button"
value="����" onclick="location.href='../welcome.asp'"></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
ȫ�������Ƶ�Ŀ�궼�ǽ���������1000��<br>
1���̲���Ҫ���������ֵ��0--1000��ģ�ÿ�����Ӽ�����ֵ1.5<br>
2��һ���������������ֵ��-500--0��ģ�ÿ�����Ӽ�����ֵ1.0<br>
3��ȫ���������������ֵ��-500���µģ�ÿ�����Ӽ�����ֵ0.5��</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
<%end if%>
<div align="center"><FONT color=#0000ff>&copy; ��Ȩ���� 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>���齭����</FONT></A></div>
</body>
</html>
