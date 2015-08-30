<%@ LANGUAGE=VBScript codepage ="936" %>
<%
function bd(zy)
zy=replace(zy,chr(39),"")
zy=replace(zy,chr(34),"")
zy=replace(zy,",","，")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="射精|奸|去死|吃屎|你妈|你娘|日你|尻|操你|干死你|王八|逼|傻B|贱人|狗娘|婊子|表子|靠你|叉你|叉死|插你|插死|干你|干死|日死|鸡巴|睾丸|死去|处女|处男|按摸|爬你达来蛋|撅你达来蛋|死你达来蛋|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|打炮|十八摸|你爷|你奶奶|你老爸|你老妈|我做你|你爸|我儿|操你|妈|逼|网管|掌门|回首当年|一线|我干|我操|你母|发铨|阿冒|阿茂|我搞|去死|法伦|法伦功|江泽民|主席|锦涛"
bad=split(badstr,"|")
for i=0 to UBound(bad)
bd=Replace(zy,bad(i),"**")
next
end function
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"yamen")=0 then 
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
	Response.End 
end if 

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
name=sjjh_name
diqu=bd(trim(request.form("diqu")))
nianling=bd(trim(request.form("nianling")))
email=bd(trim(request.form("email")))
oicq=bd(trim(request.form("oicq")))
touxiang=bd(trim(request.form("touxiang")))
qianming=bd(trim(request.form("qianming")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.Execute "update 用户 set 地区='"&diqu&"',年龄='"&nianling&"',信箱='"&email&"',Oicq='"&oicq&"',头像='"&touxiang&"',签名='"&qianming&"' where 姓名='"&Name&"'"
conn.close
set conn=nothing
%>
<html>
<title>江湖资料修改</title>
<body background="../JHIMG/BK_HC3W.GIF">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
<td bgcolor=#cccccc height="226">
<div align="center">
<p>&nbsp;</p>
<p><font color="#000000" size="3">您成功修改资料!</font><font color="#FF3333" size="3"><br>
</font></p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#000000"><br>
如果修改有错误，请重新修<br>
改即可！</font></p>
<p align=center>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</html>
