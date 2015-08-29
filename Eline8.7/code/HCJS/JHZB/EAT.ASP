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
action=request.querystring("action")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 内力,体力 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
mynl=rs("内力")
mytl=rs("体力")
rs.close
rs.open "SELECT * FROM w WHERE b=" & sjjh_id & " and i>0 and c='药品'  order by l ",conn
%>
<html>
<head>
<title>我的包袱</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body background="back.gif">
<table width="453" border="0" cellspacing="0" cellpadding="0" height="224" align="left">
  <tr> 
    <td width="65" rowspan="3" valign="top">
      <p>
        <input onClick="javascript:window.document.location.href='index.asp'" title=装备一览 type=button value="装备一览" name="button7">
        <br>
        <input onClick="javascript:window.document.location.href='daojun.asp'" title=装备头盔 type=button value="装备头盔" name="button">
        <br>
        <input onClick="javascript:window.document.location.href='eat.asp'" title=吃用药物 type=button value="吃用药物" name="button62">
        <br>姓名:<%=sjjh_name%>
        内力:<br><font color=red><%=mynl%></font><br>
      体力:<br><font color=red><%=mytl%></font>
      </p>
    </td>
    <td valign="top" rowspan="3" width="388"> 
      <div align="center"><img src="dao.gif" width="40" height="15">吃用药品一览<img src="dao1.gif" width="40" height="15"> 
        <font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div>
      <div align="center"> 注意:选择服用，将服用所有的药品！<br>
        <br>
        <table border="1" align="center" width="352" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000">
          <tr align="center"> 
            <td nowrap width="65" height="15"> 
              <div align="center"><font color="#000000">物品名</font></div>
            </td>
            <td nowrap width="30" height="15"> 
              <div align="center"><font color="#000000">数量 </font> </div>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">加内力</font></div>
            </td>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">加体力</font></div>
            </td>
            <td colspan="2" nowrap height="15"> 
              <div align="center"><font color="#000000">价值</font></div>
            </td>
            <td width="67" height="15"> 
              <div align="center"><font color="#000000">使用数量</font></div>
            </td>
            <td width="71" height="15"> 
              <div align="center"><font color="#000000">操作</font></div>
            </td>
          </tr>
          <%
do while not rs.eof
%>
          <tr> 
            <form method=POST action='wupin1.asp?action=<%=rs("a")%>&name=<%=sjjh_name%>'>
              <td width="65" height="10">
                <div align="center"><%=rs("a")%> </div>
              </td>
              <td width="30" height="10"> 
                <div align="center"><font color="#FF0000"><%=rs("i")%> </font> 
                </div>
              <td width="36" height="10"> 
                <div align="center"><font color="#0000FF"><%=rs("e")%></font> 
                </div>
              </td>
              <td width="36" height="10"> 
                <div align="center"><font color="#0000FF"><%=rs("f")%> </font> 
                </div>
              </td>
              <td colspan="2" height="10"> 
                <div align="center"><font color="#FF0000"><%=rs("l")%> </font> 
                </div>
              </td>
              <td width="67" height="10"> 
                <div align="center"> 
                  <select name="wpsl">
                    <option value="1" selected>1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="10">10</option>
                    <option value="15">15</option>
                    <option value="20">20</option>
                    <option value="25">25</option>
                    <option value="50">50</option>
                  </select>
                </div>
              </td>
              <td width="71" height="10"> 
                <div align="center"> 
                  <input type="SUBMIT" name="Submit2"  value="进行">
                </div>
              </td>
            </form>
          </tr>
          <%
rs.movenext
loop
%>
        </table>
        <br>
        <%
rs.close
rs.open "SELECT * FROM x WHERE b=" & sjjh_id & " and c>0  order by e ",conn
%>
        修练药品列表<br>
        <table border="1" align="center" width="350" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000">
          <tr align="center"> 
            <td nowrap width="65" height="15"> 
              <div align="center"><font color="#000000">物品名</font></div>
            </td>
            <td nowrap width="30" height="15"> 
              <div align="center"><font color="#000000">数量 </font> </div>
            <td nowrap width="36" height="15"> 
              <div align="center"><font color="#000000">功效</font></div>
            </td>
            <td colspan="2" nowrap height="15"> 
              <div align="center"><font color="#000000">价值</font></div>
            </td>
            <td width="66" height="15"> 
              <div align="center"><font color="#000000">使用数量</font></div>
            </td>
            <td width="70" height="15"> 
              <div align="center"><font color="#000000">操作</font></div>
            </td>
          </tr>
          <%
do while not rs.eof
%>
          <tr> 
            <form method=POST action='xleat.asp?action=<%=rs("a")%>&name=<%=name%>'>
              <td width="65" height="17">
                <div align="center"><%=rs("a")%> </div>
              </td>
              <td width="30" height="17"> 
                <div align="center"><font color="#FF0000"><%=rs("c")%> </font> 
                </div>
              <td width="36" height="17"> 
                <div align="center"><font color="#0000FF"><%=rs("d")%> </font> 
                </div>
              </td>
              <td colspan="2" height="17"> 
                <div align="center"><font color="#FF0000"><%=rs("e")%> </font> 
                </div>
              </td>
              <td width="66" height="17"> 
                <div align="center"> 
                  <select name="xlsl">
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
              <td width="70" height="17"> 
                <div align="center"> 
                  <input type="SUBMIT" name="Submit22"  value="进行">
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
        <br>
        <input onClick="JavaScript:window.close()" title=关闭窗口 type=button value="关闭窗口" name="button2">
      </div>
    </td>
  </tr>
  <tr> </tr>
  <tr> </tr>
</table>
</body>
</html>
