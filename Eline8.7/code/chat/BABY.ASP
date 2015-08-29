<%@ LANGUAGE=VBScript codepage ="936" %><%
sjjh_name=Session("sjjh_name")

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>婴儿喂养-www.51eline.com</title>
<style type="text/css">
<!--
body {  font-size: 9pt}
a:link {  text-decoration: none}
a:hover {  text-decoration: underline overline blink; font-weight: bold; cursor: crosshair; font-family: "楷体_GB2312"}
a:active {  text-decoration: line-through}
-->
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center">
<p>此项功能保留，有待开发！<br>
<br>
<font color="#FFFFFF">婴儿</font><font size="3"><br>
<br>
</font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3">【</font></span></font><font size="3" color="#000000">
喂 养</font><font color="#FFFFFF"><font size="3"> 】</font></font><font size="3"><br>
</font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> <br>
【</font></span></font><font size="3" color="#000000"> 拥 抱 </font><font color="#FFFFFF"><font size="3">】</font></font><br>
<br>
    <font color="#FFFFFF"><span style='font-size:9pt'><font size="3">2323【</font></span></font><font size="3" color="#000000"> 
    游 戏 </font><font color="#FFFFFF"><font size="3">】</font></font></p>
<p> <font color="#FFFFFF"><span style='font-size:9pt'><font size="3">【</font></span></font><font size="3" color="#000000">
学 习</font><font color="#FFFFFF"><font size="3"> 】</font></font></p>
<p>这是我们将要推出的新功能，成家的夫妻一个月后可以选择婴儿功能，将由计算机给夫妇设计一个婴儿由夫妻共同喂养，如果离婚，婴儿将会进入孤儿院，具体细节还有侍开发！</p>
</div>
</body>
</html>
