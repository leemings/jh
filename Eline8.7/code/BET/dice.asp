<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select  * from �û� where ����='"&sjjh_name&"'",conn
ying=rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>�����ӡ�һ�������wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</HEAD>

<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center"><font size="+3"><font class=c>������</font></font><br>
</div>
<div align=center>
<p>- ����������ע�� <b><font color="#0000FF">10</font> ��</b> -<br>
�����ע������ <font color="#CC0000"><b><font color="#0000FF">3000</font></b></font><b>
�� </b></p>
<form method="POST" action="dicedispose.asp">
<table border=1 cellspacing=0 cellpadding=0 align="center" width="307" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" bordercolorlight="#000000">
<tr bgcolor="#336633">
<td width="337" height="25"><font size="2" class=c><b>&nbsp;&nbsp;����ע</b></font></td>
</tr>
<tr bgcolor=#CCCCCC>
<td align=center width="337"><font color="#008000" size="2">������һ�������� <b><%=yin%>
��</b> ������Ϊ��ע</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337"><font color="#000000" size="2">��Ҫ��ע��
<input type="text" name="splosh" size="10" value="0">
&nbsp;��</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337">
<input type="submit" value="��ע��������" name="B1" >
<input type="reset" value="��Ҫ����һ�£���" name="B2" >
</td>
</tr>
</table>
</form>

<p align="center">���棺ÿ�ζ�Ǯ���������ֵ��� <font color="#FF0000">1 </font>��<br>
<br>
</p>
<p align="center"><%=Application("sjjh_chatroomname")%>��Ȩ����</p>
</div>

</BODY>
</HTML>
