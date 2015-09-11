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
<title><%=Application("aqjh_chatroomname")%>客栈</title></head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center" style="font-size:16;color:yellow"><font color="#000000">欢迎<%=aqjh_name%>光临本客栈休息
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
          <td align="center">你想要休息：
<select name="time" size="1">
              <option value="0">0小时 
              <option value="1" selected>1小时 
              <option value="2">2小时 
              <option value="3">3小时 
              <option value="4">4小时 
              <option value="5">5小时 
              <option value="6">6小时 
              <option value="7">7小时 
              <option value="8">8小时 
              <option value="9">9小时 
              <option value="10">10小时 
              <option value="11">11小时 
              <option value="12">12小时 
              <option value="13">13小时 
              <option value="14">14小时 
              <option value="15">15小时 
              <option value="16">16小时 
              <option value="17">17小时 
              <option value="18">18小时 
              <option value="19">19小时 
              <option value="20">20小时 
              <option value="21">21小时 
              <option value="22">22小时 
              <option value="23">23小时 
            </select></td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="确定"> <input
type="reset" value="取消"></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
1、在客栈中休息可以保护帐号，并可以增加内力和生命值；<br>
2、每1小时收服务费10两！增加生命10，内力10；<br>
3、请准确计算你下次使用帐号的时间！</td>
</tr>
</form>
</table>
</table>
<div align="center"><FONT color=#0000ff>&copy; 版权所有 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>爱情江湖网</FONT></A></div>
</body></html>