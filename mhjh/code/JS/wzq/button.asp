<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=017"
Set cn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
datestr=Application("yx8_mhjh_connstr")
cn.open datestr
sql="select * from game where username='"&username&"'"
set rs=cn.Execute (sql)
if rs.bOF or rs.eOF then
sql="insert into game(username,win) values ('" & username & "',0)"
set rs=cn.Execute (sql)
Response.Write "<script language=javascript>"
Response.Write "parent.top.location.href='index.htm';"
Response.Write "</script>"
Response.End
end if
sql="select * from game where username='"&username&"'"
set rs=cn.Execute (sql)
%>

<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=JavaScript>
<!--
newSize=0;
function init() {}
function resetGame()
{window.location="about:blank"}

//-->
</SCRIPT>

<META content="MSHTML 5.50.4807.2300" name=GENERATOR>

<title>������</title>
<link rel="stylesheet" href="../../STYLE.CSS">
</HEAD>
<BODY bgcolor="#FFFFFF" topmargin="0">
<FORM name=form1>
<CENTER>
<CENTER>
<TABLE cellSpacing=0 cellPadding=3 border=1 align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="471">
<TD align="center">���̴�С
<SELECT onchange=top.boardSize=parseInt(this.options[this.selectedIndex].value);top.init()>
<OPTION value=10>10</OPTION>
<OPTION value=11>11</OPTION>
<OPTION value=12>12</OPTION>
<OPTION value=13>13</OPTION>
<OPTION
value=14>14</OPTION>
<OPTION value=15 selected>ѡ��</OPTION>
<OPTION
value=16>16</OPTION>
<OPTION value=17>17</OPTION>
<OPTION
value=18>18</OPTION>
<OPTION value=19>19</OPTION>
<OPTION
value=20>20</OPTION>
</SELECT>
<FONT face=Arial,Helvetica,sans-serif size=2>
<INPUT onclick="setTimeout('top.resetGame()',100)" type=button value=������Ϸ name=newgame>
</FONT> </TD>
<TD align="center">�� �ң�<%=username%> �� ����<%=rs("win")%></TD>
</TR>
</TABLE>
<font color="#FF0033">������</font>�������а�
<table border="0" cellspacing="3" cellpadding="0" align="center">
<tr>
<%
sql="select top 10 username,win from game where win>0 order by win desc"
set rs=cn.Execute (sql)
do while not rs.eof
%>
<td align="center"><%=rs("username")%> &nbsp;<%=rs("win")%> &nbsp;</td>
<%
rs.movenext
loop
rs.Close
set rs=nothing
cn.Close
set cn=nothing
%>
</tr>
</table>
</CENTER>
</center>
</FORM>
</BODY>
</HTML>