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
rs.open "select 会员金卡,充值卡 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("会员金卡")
taiqiudian=rs("充值卡")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>金卡充值</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="bg.gif">
<div align="center"><font color="#000000"><b>欢迎<font color=blue><%=sjjh_name%></font>光临金卡充值！</b></font></div>
<br>
        <table width="300" align="center">
          <tr> 
            <td valign="top" height="8" > 
              <div align="center">
<%if hyrmb>0 then%>
  <form method="POST" action="rmbzjin.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><br>
  </div>
</form>
<%else%>
<br><br>
<%end if%>
<%if taiqiudian>0 then%>     
  <form method="POST" action="czkjk.asp" name="aff" onsubmit="bbb.disabled=1">     
  <div align="center">充值卡：<%=taiqiudian%>张，充值卡充值金卡<%=int(taiqiudian*5)%>张<br>    
    <input type="text" name="input1" size="10" maxlength="5">    
    <input type="submit" value="转换" name="bbb" class="p9" >    
  </div>    
</form>    
<%else%>    
充值卡为0不能充值<br><br>    
<%end if%>    
 <b><font color="#000000">会员金卡充值</font></b>    
                </div>   
            </td>   
          </tr>   
          <tr>    
            <td valign="top" >    
              <p>在这里你可以将充值卡，充值会员金卡！<br>转换：1张充值卡/5张金卡!</p> 
            </td> 
          </tr> 
        </table> 
<div align="center"><b><font color="#0000FF"><font color="#00FF66"><%=Application("sjjh_chatroomname")%></font></font></b> 
</div> 
</body> 
</html>