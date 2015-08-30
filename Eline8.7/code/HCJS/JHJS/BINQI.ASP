<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%><html>
<head>
<title><%=Application("sjjh_chatroomname")%>道具店♀wWw.happyjh.com♀</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/dj.gif" width="101" height="304"></td>
<td colspan="2" valign="top" align="center"> [ <font color="#009900">江 湖 道
具 店</font> ]<br>
<br>
<table width="600" border="0" cellspacing="2" cellpadding="0">
<tr>
<td width="17%" nowrap>
<div align="center"><a href="BINQI.asp"><font color="#6666FF"><b>手持刀剑</b></font></a>
</div>
</td>
<td width="15%" nowrap>
<div align="center"><a
href="anqi.asp"><font color="#6666FF"><b>江湖暗器</b></font></a></div>
</td>
</tr>
</table>
<br>
      <table border="1" align="center" width="561" cellpadding="0" cellspacing="0" bordercolor="#000000">
        <tr align="center"> 
          <td width="85"><font color="#FF6600">装 备 名</font></td>
          <td width="44"><font color="#FF6600">类 型 </font> 
          <td width="263"><font color="#FF6600">功 能</font></td>
          <td width="58"><font color="#FF6600">售 价</font></td>
          <td width="46"><font color="#FF6600">操 作</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where b<>'暗器' and b<>'卡片' and b<>'药品' and b<>'毒药' and b<>'鲜花'order by b,h ",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <form method=POST action='binqi1.asp?id=<%=rs("id")%>'>
            <td width="85">
<div align="center"><%=rs("a")%></div></td>
            <td width="44">
<div align="center"><%=rs("b")%></div></td>
            <td width="263">
<div align="center">加防御<%=abs(rs("g"))%>，加攻击<%=abs(rs("f"))%></div></td>
            <td width="58"> 
              <div align="center"><%=rs("h")%>两</div></td>
            <td width="46"> 
              <div align="center">
              <input type="SUBMIT" name="Submit" value="进行"></div>
            </td>
          </form>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
      </table>
</td>
</tr>
</table>
<br>
</body>
</html>
