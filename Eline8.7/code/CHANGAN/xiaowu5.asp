<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
sex=rs("�Ա�")
yin=rs("����")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<!--#include file="data1.asp"--><%
sql="select * from ���� where ����='" & my & "' or ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('����û�������أ�');location.href = 'fangwu.asp';}</script>"
elseif rs("״̬")<>"����" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('���ķ����ǲ����е�״��ѽ��');location.href = '../welcome.asp';}</script>"
else
lx=rs("����")
%>
<html>
<head>
<title>�� �� С ��</title>
<style>
p{font-size:9pt; color:#000000}
td,select,input{font-size:9pt; color:#000000; height:14pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body bgcolor=#990099>
<center>
<p align="center"><b><font style="font-size: 9pt" color="#ffee00">��ӭ<%=my%>�ص��Լ���С��</font></b><br><br>
<TABLE width='95%' ALIGN=center CELLSPACING=2 BORDER=2 CELLPADDING=5 BGCOLOR='#90c088'><tr><td>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>
<tr bgcolor="#C4DEFF">
<td align="center" width=10%><a href="xiaowu.asp"><font color="#000000"><%if lx="һ�㷿��" or lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>����</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu1.asp"><font color="#000000"><%if lx="һ�㷿��" or lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>����</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu4.asp"><font color="#000000"><%if lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>������</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu5.asp"><font color="#000000"><%if lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>С���</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu6.asp"><font color="#000000"><%if lx="��԰��" or lx="��������" then%>��԰</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu7.asp"><font color="#000000"><%if lx="��������" then%>��Ӿ��</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu8.asp"><font color="#000000"><%if lx="��������" then%>����</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu9.asp"><font color="#000000"><%if lx="��԰��" or lx="��������" then%>�鷿</font></a><%end if%></td>
</tr></table></TABLE>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="281">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="Ů" then%><img src="image/logonan.jpg"><%end if%><%if sex="��" then%><img src="image/logonv.jpg"><%end if%></td>                 
  <td height="1" width="380" bgcolor="#C4DEFF" rowspan="10"> 
    <%if lx="��������" then%>
    <img src="image/35028_2.jpg" width="436" height="279"> 
    <%end if%>
  </td>                 
<td height="13" width="159" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">�������</font><font color="#ff0000"><%if sex="��" then%><%=rs("����")%><%end if%><%if sex="Ů" then%><%=rs("��������")%><%end if%></font>��</p>
</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="��" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="Ů" then%><form method="post" action="putmoney.asp"><%end if%>
��Ҫ��<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>��</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="ȷ��"><INPUT TYPE="reset" VALUE="��д"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="��" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="Ů" then%><form method="post" action="putmoney.asp"><%end if%>
��Ҫȡ<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>��</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="ȷ��"><INPUT TYPE="reset" VALUE="��д"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">���ڵ��������ð�Ǯ</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">���Ƿ��ڼ���Ƚϰ�</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">ȫЩ��</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">��</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">��</td>                 
</tr>
</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">���������</font></p>
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">��ǰλ�ã����</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../welcome.asp"><font color="#cc6600">����</font></a></p>
</td>
<td height="164" width="157" bgcolor="#C4DEFF">
<font color="#ff0000">�����ϵ�������<%=yin%></font>
</td>                                                   
</tr>
</table>
<br><br><br><br><br><!--#include file="copy.inc"-->
</body></html>
<%end if
set rs=nothing	
	conn.close
	set conn=nothing
%>

