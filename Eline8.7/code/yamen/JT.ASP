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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sql="Select  * from 用户 where 姓名='"&sjjh_name&"'"
set rs=conn.Execute(sql)
if rs("道德")>-5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你还不是精神病患者你跑来这里干嘛？');}</script>"
	Response.End
end if
if rs("allvalue")>=100 then
sql="update 用户 set 通缉=False,道德=100,银两=0,存款=int(存款/2),allvalue=allvalue-100 where  姓名='"&sjjh_name&"'"
conn.Execute(sql)
else
sql="update 用户 set 通缉=False,道德=100,银两=0,存款=int(存款/2),allvalue=0 where  姓名='"&sjjh_name&"'"
conn.Execute(sql)
end if

%>
<html>

<head>
<title>江湖大教堂</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: 宋体 }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body bgcolor="#BABABA" topmargin="0" leftmargin="0">
<div align="left">
<p align="center"><font size="+2" color="#FF0000">江湖教堂</font><br>
</p>
<p align="left"><font size="+1">朋友你罪孽深重，我们来让你重生吧。。。。来感化你的灵魂：天灵灵，地灵灵各路大仙听我令，让这位大位从恶梦中醒来吧！啊门！！！</font><br>
</p>
<p align="center">&nbsp;</p>
<p align="left"><font color="#0000FF" size="+1">您的道德：+100，银两全部捐献，存款花了1/2 用于作法费用，积分减100点你终于可以重生了！</font>
<br>
</p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>

</html>
