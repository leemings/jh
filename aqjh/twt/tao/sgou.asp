<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<!--#include file="data.asp"-->
<%
sql="select * from �Խ��� where ӵ����='" & aqjh_name & "' and (��>0 or ��>0 or ͭ>0)"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "�㻹û�п���!"
  window.close()
</script>
<%
response.end
end if
 jin=rs("��")*10
 yin=rs("��")*6
 tong=rs("ͭ")*3
 money=jin+yin+tong
 conn.execute("update �Խ��� set ��=0,��=0,ͭ=0 where ӵ����='" & aqjh_name & "'")
 set rs=nothing
 conn.close
 set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
 conn.execute("update �û� set ����=����+" & money & " where ����='"&aqjh_name&"'")
 conn.close
 set conn=nothing
%>
<html>
<head>
<title>��ʯ�չ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "����"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<br>
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFCCCC"><b><%=aqjh_name%></b></font><b><font color="#FFFFFF">������Ͽ�ʯȫ���ˣ��������</font><font color="#FFCCCC"><%=money%></font><font color="#FFFFFF">��</font></b></td>
  </tr>
  <tr align="center" bgcolor="#376D95">
    <td colspan="2" height="26">
      <input type="submit" name="Submit" value="�ر�" onclick="window.close()">
    </td>
  </tr></table></body></html>