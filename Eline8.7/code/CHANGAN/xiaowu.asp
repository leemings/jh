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
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('����û�������أ�');location.href = 'fangwu.asp';}</script>"
elseif rs("״̬")<>"����" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���ķ����ǲ����е�״��ѽ��');location.href = '../welcome.asp';}</script>"
else
lx=rs("����")
%>
<html>
<head>
<title>�� �� С �ݡ�һ�������wWw.51eline.com��</title>
<style>
p{font-size:9pt; color:#ffee00}
td,select,input{font-size:9pt; color:#000000; height:9pt}
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
<td align="center" width=10%><a href="xiaowu2.asp"><font color="#000000"><%if lx="һ�㷿��" or lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>������</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu3.asp"><font color="#000000"><%if lx="�߼���Ԣ" or lx="��԰��" or lx="��������" then%>����</font></a><%end if%></td>
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
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="10"><%if lx="һ�㷿��" then%><img src="image/kt-01.jpg"><%end if%><%if lx="�߼���Ԣ" then%><img src="image/kt-02.jpg"><%end if%><%if lx="��԰��" then%><img src="image/kt-03.jpg"><%end if%><%if lx="��������" then%><img src="image/kt-04.jpg"><%end if%> </td>                 

</tr>
<tr>
  <td height="27" width="159" bgcolor="#C4DEFF" > 
    <p align="center"><a href="moshu.asp"><font color="#000000">ѧϰħ��</font></a></p> 
��</td>                 
</tr>
<tr>
<td height="20" width="159" bgcolor="#C4DEFF">
<p align="center"><a href="shangwang.asp"><font color="#000000">�ϻ�����</font></a></p> 
��</td>                 
</tr>
<tr>
<td height="23" width="159" bgcolor="#C4DEFF">
<p align="center"><a href="kadianshi.asp"><font color="#000000">�����Ӱ�</font></a></p> ��</td>                 
</tr>
<tr>
<td height="22" width="159" bgcolor="#C4DEFF">��
<p align="center"><a href="hekafei.asp"><font color="#000000">��������</font></a></p>                 
</tr>
<tr>
              
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
set conn=nothing
%>