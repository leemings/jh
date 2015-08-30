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
rs.open "select 金币,银币,会员等级 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("金币")
taiqiudian=rs("银币")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>娱乐金币♀wWw.happyjh.com♀</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../chat/back.jpg">
<div align="center"><b><font color="#800000">欢迎<%=sjjh_name%>大侠<img border="0" src="../chat/img/bt.gif">光临娱乐金币转换处！</font></b></div>
<br>
        <table width="300" align="center" bordercolor="#339966" border="1">
          <tr> 
            <td valign="top" height="8" > 
              <div align="center">
<%if hyrmb>0 then%>
  <form method="POST" action="rmbzjin6.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#339966"><img border="0" src="../chat/img/251.GIF">金币：<%=hyrmb%>个，金币转换银币<%=hyrmb*10%>个<br>
    <input type="text" name="input" size="10" maxlength="5">
    <input type="submit" value="转换" name="bb" class="p9">  
    </font>
  </div>
</form>
<font color="#339966">
<%else%>
金币为0不能转换<br><br>
<%end if%>
<%if taiqiudian>0 then%>    
</font>  
  <form method="POST" action="jinzrmb6.asp" name="aff" onsubmit="bbb.disabled=1">  
  <div align="center"><font color="#339966">银币：<%=taiqiudian%>个，银币转换金币<%=int(taiqiudian/10)%>个<br> 
    <input type="text" name="input1" size="10" maxlength="5"> 
    <input type="submit" value="转换" name="bbb" class="p9" >  
    </font> 
  </div> 
</form> 
<font color="#339966"> 
<%else%> 
银币为0不能转换<br><br> 
<%end if%> 
               娱乐金币</font>    
                </div> 
            </td> 
          </tr> 
          <tr>  
            <td valign="top" >  
              <p><font color="#800000">在这里你可以将卖豆点的钱转换成金币，金币可以进行台球娱乐活动！<br>转换：1金币/10个银币!银币10个换金币/1个</font></p>
            </td>
          </tr>
        </table>
<div align="center">
</div>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body>
</html>