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
sex=sjjh_jhdj
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
<title>�� �� С �ݡ�һ�������wWw.happyjh.com��</title>
<style>
p{font-size:9pt; color:#ffee00}
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
<p align="center"><b><font style="font-size: 9pt">��ӭ<%=my%>�ص��Լ���С��</font></b><br><br>
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
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="11"><%if sex="Ů" then%><img src="image/logonan.jpg"><%end if%><%if sex="��" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="11"><%if lx="һ�㷿��" then%><img src="image/ws-01.jpg"><%end if%><%if lx="�߼���Ԣ" then%><img src="image/ws-02.jpg"><%end if%><%if lx="��԰��" then%><img src="image/ws-03.jpg"><%end if%><%if lx="��������" then%><img src="image/ws-04.jpg"><%end if%> </td>                 
<td height="1" width="159" bgcolor="#C4DEFF">
<form method=POST action='pub1.asp'>
����Ҫ��Ϣ��</td>                 
</tr>
<tr><td height="1" width="159" bgcolor="#C4DEFF">
<select name=date size=1>
<option value="0">����
<option value="1">һ��
<option value="2">����
<option value="3">����
</select>
<select name=time size=1> 
<option value="0">0Сʱ
<option value="1">1Сʱ
<option value="2">2Сʱ
<option value="3">3Сʱ
<option value="4">4Сʱ
<option value="5">5Сʱ
<option value="6">6Сʱ
<option value="7">7Сʱ
<option value="8">8Сʱ
<option value="9">9Сʱ
<option value="10">10Сʱ
<option value="11">11Сʱ
<option value="12">12Сʱ
<option value="13">13Сʱ
<option value="14">14Сʱ
<option value="15">15Сʱ
<option value="16">16Сʱ
<option value="17">17Сʱ
<option value="18">18Сʱ
<option value="19">19Сʱ
<option value="20">20Сʱ
<option value="21">21Сʱ
<option value="22">22Сʱ
<option value="23">23Сʱ
</select>
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
<INPUT TYPE="submit" VALUE="ȷ��"><INPUT TYPE="reset" VALUE="��д">
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">����������ϢÿСʱ</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">������1000����1000</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">����Ҫ������</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">��</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
    <p align="center">&nbsp;</p>
  </td>                 
</tr></form></table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
<p align="center"><a href="../XIAOHAI/feedsheep.asp" target="_blank"><font color="#000000">���Ϻ���</font></a></p> 
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">��ǰλ�ã�����</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../welcome.asp"><font color="#cc6600">����</font></a></p>
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