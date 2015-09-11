<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML>
<HEAD>
<title>江湖逃婚</title>
<link rel="stylesheet" href="../../css.css">
</HEAD>
<body bgcolor="#8EB4D9" background="../../bg.gif">
<p align=center> <b> <font color="#000000" size="4">我要逃婚</font></b></p>

<p align="center"><img border="0" src="no_27.gif"></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="宋体" size="2">我是超时空博士，我现在是回到这个时代给你们解答</font></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="宋体" size="2">问题，如果你不想要现在的配偶，或者你很少见你的</font></p>
<p align="center"><font lang="ZH-CN" color="#FFFFFF" face="宋体" size="2">配偶，觉得寂寞难耐了。那么，我可以帮你<a href="panjue1.asp">实现逃婚</a></font></p>
<p align="center">　</p>
</body></HTML> 