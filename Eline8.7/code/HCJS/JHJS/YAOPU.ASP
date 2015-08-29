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
lx=request.querystring("lx")
if lx="" then lx="药品"
if lx<>"药品" and lx<>"大赛" then
	Response.Write "<script language=JavaScript>{alert('类型出错！');}</script>"
	Response.End
end if
if instr(lx,"or")<>0 or instr(lx,"'")<>0 or instr(lx,"=")<>0 or instr(lx,">")<>0 or instr(lx,"<")<>0 then
	Response.Write "<script language=JavaScript>{alert('想黑我没门！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
%>
<html>
<head>
<title>〖<%=Application("sjjh_chatroomname")%>〗</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/yq.gif" width="101" height="304"></td>
    <td colspan="2" valign="top" align="center"> 
      <p align="center">[ <a href="YAOPU.ASP?lx=药品"><font color="#6633FF">药 品 专 卖 店</font></a> 
        ] [<a href="YAOPU.ASP?lx=大赛"><font color="#6633FF">夺宝大赛专用药品</font></a>] [<a href="DUYAO.asp"><font color="#6633FF">毒 药 专 卖 店</font></a>]<br>
        <font color="#000000">经营者将获得药价的%5为利润！</font></p>
      <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width="461">
        <tr> 
          <td width="50" height="13"> 
            <div align="center">补 药</div>
          </td>
          <td width="234" height="13"> 
            <div align="center">功 能</div>
          </td>
          <td width="59" height="13"> 
            <div align="center">售 价</div>
          </td>
          <td width="55" height="13"> 
            <div align="center">数 量</div>
          </td>
          <td width="51" height="13"> 
            <div align="center">操 作</div>
          </td>
        </tr>
        <%
rs.open "SELECT * FROM b where  b='"&lx&"' order by h",conn
do while not rs.bof and not rs.eof
%>
        <tr> 
          <form method=POST action='buyyao.asp?id=<%=rs("id")%>'>
            <td width="50"> 
              <div align="center"><%=rs("a")%></div></td>
            <td width="234"> 
              <div align="center">补内力<%=rs("d")%>，补生命<%=rs("e")%>! 经营者:<%if rs("n")<>"无" then%><a href="#" onClick="window.open('../../chat/myjb.asp?wpname=<%=rs("a")%>','myjb','scrollbars=yes,resizable=no,width=420,height=330')" title="<%=ti%>"><%=rs("n")%></a><%else%><%=rs("n")%><%end if%></div>
            </td>
            <td width="59"> 
              <div align="center"><%=rs("h")%>两</div></td>
            <td width="55"> 
              <div align="right">
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
                  <option value="30">30</option>
                  <option value="50">50</option>
                  <option value="99">99</option>
                </select>
              </div>
            </td>
            <td width="51"> 
              <div align="center"> 
                <input type="SUBMIT" name="Submit" value="进行">
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
<br></body>

</html>
