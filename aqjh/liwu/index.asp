<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../mywp.asp"-->
<!--#include file="../showchat.asp"-->

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');</script>"
	Response.End
end if



%>
<html>
<head>
<title>礼物发放</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body text="#000000" background="../chat/bg.gif">
<div align="center">
  <tr>
  
    <td width="80" align="center"> 
       <input type=button value='新人礼物' onClick="window.open('../liwu/xrlw.asp','f3')" >
</div>
<div align="center">
<input type=button value='周五礼物' onClick="window.open('../liwu/zwlw.asp','f3')" >
</div>
<div align="center">
<input type=button value='周六礼物' onClick="window.open('../liwu/zllw.asp','f3')" >
</div>
<div align="center">
<input type=button value='周日礼物' onClick="window.open('../liwu/zmlw.asp','f3')" >
</div>
<div align="center">
<input type=button value='月末礼物' onClick="window.open('../liwu/ymlw.asp','f3')" >
</div>
<div align="center">
<input type=button value='站长礼物' onClick="window.open('../liwu/zzlw.asp','f3')" >
</div>
<div align="center">
<input type=button value='节日礼物' onClick="window.open('../liwu/jrlw.asp','f3')" >




 </td>

  </tr>
</div>
</body>
</html>
</body>
</html>
