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
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>���ɻ����wWw.51eline.com��</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
<div align="center"><font color="#000000"><b>��ӭ<font color=blue><%=sjjh_name%></font>�������ɻ����</b></font></div>
<form method=POST action='putmpjj.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
<td align=center>��ѡ�����Դ����Ǯ����
<select name=money size=1>
<option value="1000" selected>1000</option>
<option value="10000">10000</option>
<option value="100000">100000</option>
<option value="1000000">1000000</option>
</select>
</td>
</tr>
<tr>
<td  align=center>
<input type=submit value=���˾����� name="submit">
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
���ɻ���</div>
</td>
</tr>
<tr>
<td valign="top" >
<p>���ɻ�����һ��ȡ֮������֮��������ߣ���ȡ�ý�����ڱ����ɽ���У���ִ�����Ϸ��������Ҫ���ˡ����Ա��ɵĸ����һ���Ӧ�еĻر��ġ����ڱ���֧������ӱ���ȡǮ���������Ƕ����м�¼���еģ��뵽���������в�ѯ��</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body>
</html>



