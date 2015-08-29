<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 会员金卡,金币 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("会员金卡")
taiqiudian=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>娱乐金币♀wWw.51eline.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../chat/back.jpg">
<div align="center"><font color="#000000"><b>欢迎<font color=blue><%=sjjh_name%></font>光临娱乐金币转换处！</b></font></div>
<br>
        <table width="300" align="center">
          <tr> 
            <td valign="top" height="8" > 
              <div align="center">
<%if hyrmb>0 then%>
  <form method="POST" action="rmbzjin.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center">现金：<%=hyrmb%>元，现金转换金币<%=hyrmb*10%>个<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9">
  </div>
</form>
<%else%>
会员金卡为0不能转换<br><br>
<%end if%>
<%if taiqiudian>0 then%>
  <form method="POST" action="jinzrmb.asp" name="aff" onsubmit="bbb.disabled=1"> 
  <div align="center">金币：<%=taiqiudian%>个，金币转换现金<%=int(taiqiudian/10)%>元<br>
    <input type="text" name="input1" size="10" maxlength="5">
    <input type="submit" value="转换" name="bbb" class="p9" >
  </div>
</form>
<%else%>
金币为0不能转换<br><br>
<%end if%>
               娱乐金币
                </div>
            </td>
          </tr>
          <tr> 
            <td valign="top" > 
              <p>在这里你可以将卖豆点的钱转换成金币，金币可以进行台球娱乐活动！<br>转换：1元/10点金币!同样您也可以将金币转换成现金购买卡片！</p>
            </td>
          </tr>
        </table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body>
</html>