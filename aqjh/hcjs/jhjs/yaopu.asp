<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>〖<%=Application("aqjh_chatroomname")%>药品药专卖店〗</title>
<link rel="stylesheet" href="../../css.css">
</head>

<body topmargin="6"
leftmargin="0" bgcolor="#FFFFFF" background="../../bg.gif">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td width="101" align="center" valign="top"><img src="../../images/yq.gif" width="101" height="304"></td>
    <td colspan="2" valign="top" align="center"> 
      <p align="center"> <font color="#6633FF">〖<%=Application("aqjh_chatroomname")%>药品药专卖店〗</font><br></p>
      <font size="+1"><a href="DUYAO.asp"><b>去买毒药去</b></a></font><br>
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
rs.open "SELECT * FROM b where  b='药品' order by h",conn
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
                  <option value="10">10</option>
                  <option value="99">99</option>
                  <option value="999">999</option>
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
</table><br></body></html>