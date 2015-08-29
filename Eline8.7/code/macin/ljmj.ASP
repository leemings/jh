<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
'if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.Open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs("体力")<500 then
Response.Write "<script Language=Javascript>alert('提示：打麻将扣体力500！');window.close();</script>"
	response.end
end if
rs.Close 
conn.Close 
%>

<head>
<link rel="stylesheet" href="../style4.css">
<title>麻将♀wWw.51eline.com♀</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false background="../images/8.jpg" leftmargin="0" topmargin="0">
<!-- URL's used in the movie-->
<!-- text used in the movie-->
<!--Loading... 27k “斡巳□□髯沂 □暗坛蛐枭萍 Next Level Line High Score Game Over --><OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
 codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
 WIDTH=550 HEIGHT=400>
 <PARAM NAME=movie VALUE="bjmj.swf"> <PARAM NAME=menu VALUE=false> <PARAM NAME=quality VALUE=medium> <PARAM NAME=bgcolor VALUE=#000000> <EMBED src="bjmj.swf" menu=false quality=medium bgcolor=#000000  WIDTH=550 HEIGHT=400 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
</OBJECT>
</BODY>
</html>





