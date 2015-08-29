<%@ LANGUAGE=VBScript codepage ="936" %><%

' --- 各项运程内容设定
luck1 = "<img src=img/01.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>大吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>似乎你今天的运气很旺盛!<br>有什么想做的事便去尝试吧,<br>或许会有成果!<br>又或者赶快去买个彩票吧!"
luck2 = "<img src=img/02.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>中吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>恭喜你呀!<br>今天应该事事如意,<br>或许坏天气也会转晴呢!<br>不过凡事都是脚踏实地好了,<br>这只是种小玩意!"
luck3 = "<img src=img/03.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFF>小吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>心情也算不错了吧!<br>虽然有时会遇到失败,<br>但切勿放弃啊!"
luck4 = "<img src=img/04.gif border=0></td></tr><tr><td align=center height=15 bgcolor#c00000><font color=#FFFFF>小凶</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>也许今天不适宜干大事,<br>睡个觉就好了!"
luck5 = "<img src=img/05.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c0000><font color=#FFFFFF>中凶</font></td></tr><tr><td bgcolor#FFdbdb align=center valign=top height=130>噢不要在乎这些无聊玩意!<br>事情好与坏是由你决定的,<br>别灰心呀!"


randomize
x = int(rnd * 100)
' 确率 20%
if x < 20 then 
 luck = luck1
' 确率 20%
elseif x < 40 then
 luck = luck2
' 确率 30%
elseif x < 70 then
 luck = luck3
' 确率 20%
elseif x < 90 then
 luck = luck4
' 确率 10%
else luck = luck5
end if
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=gb2312">
<title>今日运程♀wWw.51eline.com♀</title>
<STYLE TYPE="text/css">
<!--
tr, td,body,th    {font-size: 9pt}
a:link    {font-size: 9pt; text-decoration:none; color:<%=link%> }
a:visited {font-size: 9pt; text-decoration:none; color:<%=vlink%> }
a:active  {font-size: 9pt; text-decoration:none; color:<%=alink%> }
a:hover   {font-size: 9pt; text-decoration:underline; color:<%=alink%> }
input,select,textarea     {font-size: 9pt; border: 1 solid black}
-->
</STYLE>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<div align="left">

<table border="1" cellspacing="1" width="157"
bordercolor="#FFADAD">
    <tr>
        <td height=142 align=center valign=middle><%=luck%>
        </td>
    </tr>
</table>
</div>
</body>
</html>
