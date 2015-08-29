<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>暗器店〗</title>
<link rel="stylesheet" href="../../css.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/dj.gif" width="101" height="304"></td>
    <td colspan="2" valign="top" align="center"><font color="#6633FF">〖<%=Application("sjjh_chatroomname")%>暗器店〗♀wWw.51eline.com♀</font><br>
<br>
      <br>
      <table border="1" align="center" width="546" bordercolor="#000000" cellspacing="0">
        <tr align="center"> 
          <td width="69"><font color="#FF6600">装备名</font></td>
          <td width="58"><font color="#FF6600">类型 </font> 
          <td width="244"><font color="#FF6600">功能</font></td>
          <td width="57"><font color="#FF6600">售价</font></td>
          <td colspan="2"><font color="#FF6600">数量</font></td>
          <td width="46"><font color="#FF6600">操作</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM b where b='暗器' order by h ",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <form method=POST action='binqi3.asp?id=<%=rs("id")%>'>
            <td width="69"><div align="center"><%=rs("a")%></div></td>
            <td width="58"><div align="center"><%=rs("b")%></div></td>
            <td width="244"> 
              <div align="center">伤内力<%=abs(rs("d"))%>，伤体力<%=abs(rs("e"))%></div>
            </td>
            <td width="57"> 
              <div align="right"><%=rs("h")%>两</div></td>
            <td colspan="2"> 
              <div align="center"> 
                <select name="sl">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div>
            </td>
            <td width="46"> 
              <div align="center"> 
                <input type="SUBMIT" name="Submit"
value="进行">
              </div>
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
</body>
</html>
