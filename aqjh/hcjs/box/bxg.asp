<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../mywp.asp"-->
<%
Response.Write "<SCRIPT LANGUAGE=javascript>if(window.name!='bxg'){this.location.href='index.asp'}</SCRIPT>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs1.open "select id,��Ա�ȼ�,w1,w2,w3,w4,w5,w6,w7,w8 from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs1.eof or rs1.bof then
	rs1.close
	set rs1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('�㲻�ǽ������ˣ�');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
myid=rs1("id")
hydj=rs1("��Ա�ȼ�")

wpzs=10
wpzs=wpzs+hydj*5	'�ó����ɴ�ż�����Ʒ
rs.open "select * from w where a="&myid&" order by d",conn,2,2
%>
<html>
<head>
<title>�ҵĴ�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
a:active {  font: 12px "����"; color: #FFFFFF; background: #000000}
a:link {  font: 12px "����"; color: #FFFFFF}
a:visited {  font: 12px "����"; color: #FFFFFF}
body {  font: 12px "����"; color: #FFFFFF}
input {  font: 12px "����"; color: #FFFFFF; text-decoration: none; background: #000000; border: none}
td {  font: 12px "����"; color: #FFFFFF; text-align: center}
-->
</style>
<link rel="stylesheet" href="STYLE.CSS" type="text/css">
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" link="#00CCFF" vlink="#00CCFF" alink="#00CCFF" oncontextmenu=self.event.returnValue=false>
<br>
<div align="center"><font size="5" color="#FF0000"><b class="3d"><font size="6">�ҵĴ�����</font></b></font><br>
<br>
<font color="#FFFF00">����ұߵ���Ʒ�����ٰ������봢���䡱��ť���ɷ��봢���䡣</font><br><br>
<a href="#" onclick="javascript:parent.bxg.document.location.reload();">���ˢ�´�����</a><br>
<table width="82%" border="1" cellpadding="0" cellspacing="0">
<tr>
<form name="bxgform" method="post" action="inthebox.asp">
<td colspan="5" height="29">
<input type="hidden" name="filedname" size="5" maxlength="5" value="">
<input type="text" readonly name="wpname" size="10" maxlength="15" class="input">
<input type="text" name="sl" size="5" maxlength="5" class="input">
<input type="submit" name="Submit2" value="���봢����" style="border: 1 solid #FFFFFF">
</td>
</form>
</tr>
<tr>
<td width="199" height="24"><font color="#FFCC66"><b>��Ʒ����</b></font></td>
<td width="117" height="24"><font color="#FFCC66"><b>�������</b></font></td>
<td width="93" height="24"><font color="#FFCC66"><b>����</b></font></td>
<td width="68" height="24"><font color="#FFCC66"><b>ȡ������</b></font></td>
<td width="150" height="24"><font color="#FFCC66"><b>����</b></font></td>
</tr>
<%jsq=0
if rs.eof or rs.bof then%>
<tr><td colspan="5" height="29">��Ĵ�������ʲôҲû��</td></tr>
<%else
do while not rs.bof and not rs.eof
	tempp=""
if jsq>=wpzs then
	yhfiled=rs("d")
	yid=rs("a")
	wid=rs("id")
	wm=rs("b")
	ws=rs("c")
	conn.execute "delete * from w where id="&wid
	tempp=add(rs1(yhfiled),wm,ws)
	conn.execute "update �û� set "&yhfiled&"='"&tempp&"' where id="&yid
else
%>
<tr>
<td width="199"><%=rs("b")%></td>
<%sslb="��"
select case rs("d")
	case "w1"
		sslb="��ҩ"
	case "w2"
		sslb="��ҩ"
	case "w3"
		sslb="װ��"
	case "w4"
		sslb="����"
	case "w5"
		sslb="��Ƭ"
	case "w6"
		sslb="��Ʒ"
	case "w7"
		sslb="�ʻ�"
	case "w8"
		sslb="��ҩ"
	case else
		sslb="������"
end select
%>
<td width="117"><%=sslb%></td>
<td width="93"><%=rs("c")%></td>
<form name="form2" method="post" action="qwupin.asp?id=<%=rs("id")%>">
<td width="68" align="center">
<input type="text" name="sl" size="5" maxlength="5" style="background-color: #800000; color: #FFFFFF; border: 1 solid #52436D">
</td>
<td width="150" align="center">
<input type="submit" name="Submit3" value="ȡ����Ʒ"  style="border: 1 solid #FFFFFF">
</td>
</form>
</tr>
<%jsq=jsq+1
end if
rs.movenext
loop%>
<tr>
<td colspan="5" height="29"><br>
        ��Ĵ����乲�����<font color="#FFFF00"><%=jsq%></font>����Ʒ<br>
<br>
<div align="left"><font color="#00FFFF">�ǻ�Ա���Դ��10����Ʒ��һ����Ա�ɴ��15����Ʒ<br>
          ������Ա�ɴ��20����Ʒ��������Ա�ɴ��25����Ʒ���ļ���Ա�ɴ��30����Ʒ<br>
          ����Ա���ں����Ʒ��������Ӧ�е�����ʱ���Զ����������зŲ��µ���Ʒ�Զ��ӵ��������ϡ�</font><br>
</div>
</td>
</tr>
<%end if
rs1.close
set rs1=nothing
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<br>
<FONT color=#0000ff>&copy; ��Ȩ���� 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A></div></body></html>