<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from �û� where ����='" & aqjh_name &"'",conn,3,3
%>
<html><head>
<link rel="stylesheet" href="../../css.css">
<title>é��--[���ֽ��������״�]</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style fprolloverstyle>A:hover {color: #FF0000; font-weight: bold}
</style>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../chat/bg.gif">
<table border="1" cellspacing="0" cellpadding="0" width="442" align="center" bordercolor="#FFFFFF">
<tr>
    <td width="438" height="335" valign="top"> 
      <div align="center"><b><font color=blue><%=aqjh_name%></font><font color="#ff0000">����ո��ˡ�����վ��<img border="0" src="../../chat/pic/dz01.gif" loop="infinite"></font></b></div>
<form method=POST action='maoceok.asp'>
<table width="536" align="center" bordercolor="#FBEFD7" border="1">
<tr>
<td height="21" width="526">
<tr>
            <td align=center height="46" width="526"><font color="#FF0000"><strong>���˷�⣺<%=rs("���")%></strong></font> 
              
</td>
</tr>
<tr>
            <td  align=center height="40" width="526"><font color="#FF0000"><strong>���ѣ�</strong></font> 
              <select name=money size=1>    
                <option value="10000000">10000000</option>
              </select>
              <input type=submit value=��Ҫ���� name="submit">    
</td>
</tr>
<tr>
<td valign="top" height="40" width="526" >
<div align="center">
<br>
<br>
<strong><font size="4" color="#FF0000">��&nbsp;&nbsp;&nbsp; ��&nbsp;&nbsp; 
��&nbsp; ��</font></strong></div>
</td>
</tr>
<tr>
            <td valign="top" width="526" > 
              <p align=center><font color="#FF0000"><strong>�Լ�������Ҫ1000����������ְҵΪ����ҹ�㡱������Ϊ</strong></font></p>
              <p align=center><font color="#FF0000"><strong>2��������������Ӵ˹��ܲ���Ϊ�������˵ķ���~~~~����</strong></font></p>
              <p align=center><font color="#FF0000"><strong>~~~~~������Ϊ�˷����ҵģ���һ���Ҳ���ְҵΪ����ҹ</strong></font></p>
              <p align=center><font color="#FF0000"><strong>�㡱�����,���˶಻���㣬������ֻ�ùԹ��������ˣ�:��</strong></font></p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>