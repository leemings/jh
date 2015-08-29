<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>管理员名单♀wWw.51eline.com♀</title>
<link rel="stylesheet" href="../css.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="190"><img src="../images/gf1.jpg" width="57" height="323"><img src="../images/gf2.jpg" width="58" height="323"><img src="../images/gf3.jpg" width="65" height="323"></td>
<td width="588" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td>
<table width="97%" align="center" cellspacing="1" border="0"
cellpadding="0" bordercolor="#000000">
<tr>
<td colspan="5">
<div align="center"><font color="#FF6600">官 府 网 管 资 料<br>
</font></div>
</td>
<tr bgcolor="#FFCC33">
<td align="center" width="19%" height="21"> 姓 名 </td>
<td align="center" width="11%" height="21"> 等 级 </td>
<td align="center" width="14%" height="21" bgcolor="#FFCC00">
状 态 </td>
<td align="center" width="56%" height="21"> 站长操作 </td>
</tr>
<tr>
<td width="19%" nowrap>
<div align="center"> </div>
</td>
<td width="11%" nowrap>
<div align="center"> </div>
</td>
<td width="14%" nowrap>
<div align="center"> </div>
</td>
<td width="56%" nowrap>
                  <div align="center">此项功能请在站长管理操作</div>
</td>
</tr>
</table>
</td>
</tr>
<tr>
          <td height="16"> <div align="left"><font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;▼<font size="2">官府人员权限与职责</font></font><font size="2"></font></div></td>
</tr>
<tr>
          <td> <ul>
              <div align="left"><font size="2">
                </font></div>
              <div align="left"> </div>
              <li><font color="#000000" size="2">官府管理员必须管理刷屏，严格管理骂人、以及违反中国有关互联网规定的对象！<br>
                </font></li>
              <li>必须时常保持在线。如果发现不遵守江湖规则、恶意扰乱秩序警告不听者，一律封杀！严重者锁其IP！</font>
                </li>
              <br>
              <li>官府身份分为6--10级，按照身份的高低拥有不同的权力。其中：<br>
              </li>
              <li>警告与踢人为基本权力<br>
              </li>
              <li>6级以上身份可以逮捕<br>
              </li>
              <li>7级以上可以抓人坐牢，10分钟不可进江湖（踢人）<br>
              </li>
              <li>8级以上可以发钱（暂时封闭这个功能，怕乱）<br>
              </li>
              <li>9级以上可以对牢里的犯人进行斩首<br>
              </li>
              <li>10级为超级管理员<br>
              </li>
              <li>发现管理员有乱用权力的，应当及时举报投诉，站长将在查清事实后进行严肃的处理！</li>
            </ul>
</td>
</tr>
</table>
</td>
</tr>
</table>
<div align="center"><br>
<font color="#000000" size="-1"><br>
【 授权：<%=Application("sjjh_user")%> 运行序列号：<%=Application("sjjh_sn")%> 】 </font></div>
</body>

</html>