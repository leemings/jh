<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("aqjh_chatroomname")%>��ջ</title></head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center" style="font-size:16;color:yellow"><font color="#000000">��ӭ<%=aqjh_name%>���ٱ���ջ��Ϣ
</font>
<table border="1" bgcolor="#FFCC99" align="center" width="399" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#BEE0FC" height="230">
<table width="358">
<tr>
<td>
<form method="POST" action="pub1.asp">
<tr>
          <td align="center">����Ҫ��Ϣ��
<select name="time" size="1">
              <option value="0">0Сʱ 
              <option value="1" selected>1Сʱ 
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
            </select></td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="ȷ��"> <input
type="reset" value="ȡ��"></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
1���ڿ�ջ����Ϣ���Ա����ʺţ���������������������ֵ��<br>
2��ÿ1Сʱ�շ����10������������10������10��<br>
3����׼ȷ�������´�ʹ���ʺŵ�ʱ�䣡</td>
</tr>
</form>
</table>
</table>
<div align="center"><FONT color=#0000ff>&copy; ��Ȩ���� 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>���齭����</FONT></A></div>
</body></html>