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
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>�¶�Ժ</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=aqjh_name%></font>���ٰ���¶�Ժ</b></font></div>
<form method=POST action='gryok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
<td align=center>�ҵİ��ģ� 
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
<input type=submit value=ף����Щ���� name="submit">
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                �¶�Ժ���</div>
</td>
</tr>
<tr>
<td valign="top" >
              <p>ս����ƶ��ʹ��������Щ����ʧȥ�˸�ĸ�ף������������Ƕ�ôϣ���õ���������ϵİ��������׳���İ��ģ�Ϊ�������Ĺ¶�����һƬ���ģ�(���װ��Ŀ���ʹ���¸����ڵ�1�����=500����)</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF">��Ȩ���С����齭����վ��</font></b></font>
</div>
</body>
</html>