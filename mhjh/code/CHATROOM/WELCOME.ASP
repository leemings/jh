<%@ LANGUAGE=VBScript.Encode%>
<!--#INCLUDE FILE="midsound.asp" -->
<%
un=Request("search")
kao=Request("do")
linkurl="mid/" & un &".mid"
%>
<html>
<head>
<link rel="stylesheet" href="css.css">
<title></title>
<bgsound src="<%=linkurl%>" loop="<%=kao%>">
</head>
<body  background="bg1.gif" style="CURSOR: url('../images/banana.ani')" oncontextmenu=self.event.returnValue=false>
<table align="center" width="490">
  <tr> 
    <td width="482" height="17">
<font color="#FF0000">自动申请会员条件降低,5000分即可申请成功;娱乐中心增加获得精力游戏项目。</font></td>
  </tr>
</table>                          
</body>                          
</html>