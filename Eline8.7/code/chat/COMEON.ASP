<%@ LANGUAGE=VBScript codepage ="936" %>
<html>
<head>
<title>回来吧 - <%=Application("sjjh_chatroomname")%></title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FF0000" text="#FFFFFF">
<p align="center"><font color="#FFFF00">★★☆☆ 快回来吧 ☆☆★★ </font></p>
<p align="center"> 再过 300 秒您就要沉到海底了<br>
快回来和大家说说话哦！*^o^*</p>
<p align="right"><font color="#FFFFFF">──【<%=Application("sjjh_chatroomname")%>】</font></p>
<bgsound src="readonly/ri.mid" loop="1">
<script Language="JavaScript">function re(){window.close();}var tid=null;tid=setTimeout('re()',300000);</script>
</body>
</html>