<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();}</script>"
	Response.End 
end if
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM W WHERE b="& sjjh_id & " and i>0 and j=false and c<>'��Ƭ' order by c ",conn
%>
<html>
<head>
<title>��Ʒ�����wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#000000" leftmargin="0" topmargin="0">
<div align="left">
<div align="center">ʱ�䣺<font color="#000000" size="-1"><%=now()%></font>������Ʒһ��(װ������Ʒ����)<font face="��Բ"><a href="javascript:this.location.reload()">ˢ��</a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="558" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="53" height="16">
<div align="center"><font color="#000000" size="-1">��Ʒ��</font></div>
</td>
<td nowrap width="30" height="16">
<div align="center"><font color="#000000">���� </font></div>
<td nowrap width="40" height="16">
<div align="center"><font color="#000000" size="-1">���� </font> </div>
<td nowrap width="41" height="16">
<div align="center"><font color="#000000">���� </font></div>
<td nowrap width="42" height="16">
<div align="center"><font color="#000000">���� </font></div>
<td nowrap width="39" height="16">
<div align="center"><font color="#000000">���� </font></div>
<td nowrap width="42" height="16">
<div align="center"><font color="#000000">���� </font></div>
<td colspan="2" nowrap height="16">
<div align="center"><font color="#000000" size="-1">��Ǯ</font></div>
</td>
<td nowrap width="56" height="16">
<div align="center"><font color="#000000">��ʽ</font></div>
</td>
<td nowrap width="48" height="16">
<div align="center"><font color="#000000">����</font></div>
</td>
<td nowrap width="54" height="16">
<div align="center"><font color="#000000">����</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';">
<form method=POST action='fs.asp?id=<%=rs("id")%>&name=<%=name%>'>
<td width="53" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("a")%>
</font> </div>
</td>
<td width="30" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("c")%></font>
</div>
<td width="40" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("i")%>
</font> </div>
<td width="41" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("e")%></font>
</div>
<td width="42" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("f")%></font>
</div>
<td width="39" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("g")%></font>
</div>
<td width="42" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("h")%></font>
</div>
<td colspan="2" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("l")%>
</font> </div>
</td>
<td height="3" width="56">
<div align="center"><font color="#000000">
<select name="wpfs">
<option value="1" selected>&nbsp;����</option>
<option value="2">&nbsp;����</option>
<option value="3">&nbsp;����</option>
<option value="4">���չ�</option>
</select>
</font></div>
</td>
<td height="3" width="48">
<div align="center"> <font color="#000000">
<select name="wpsl">
<option value="1" selected>1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>
</select>
</font></div>
</td>
<td height="3" width="54">
<div align="center"><font color="#000000">
<input type="SUBMIT" name="Submit"  value="����">
</font></div>
</td>
</form>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table width="64%" border="1" cellpadding="0" cellspacing="0">
<tr>
<td>˵��������ʱ����Զ��Լ�����Ʒ���й��������<br>
1.���ͣ����Լ��Ķ������͸������������˲���ֻ�Ƿ���������֮�䡣�����ô˽��н�����<br>
2.���ף��������������ѽ�����Ʒ���ף��������򲻵��ġ���Ǯ���㶨�����ֵ��������<br>
3.���֣��������������Ҫ�ۣ�һ��Ը����һ��Ը�򡣲������ˡ��ҡ��˴�ͷ�����ǿɲ��ܡ�<br>
4.���չ���ʲô�������Լ�͵͵��������������Ǻõط�ѽ���ǲ��ᶪ�ģ������շ�Ӵ��</td>
</tr>
</table>
</div>
</div>
</body>
</html>
