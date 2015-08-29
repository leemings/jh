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
rs.open "select 法力,金币 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("法力")
taiqiudian=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>法力转换♀wWw.51eline.com♀</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../chat/back.jpg">
<div align="center"><b><font color="#008080">欢迎大侠<img border="0" src="../chat/img/mob_s_20.gif" width="100" height="100">光临法力转换处！</font></b></div>
<br>
        <table width="300" align="center" bordercolor="#008080" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#800000" > 
              <div align="center">
<%if hyrmb>0 then%>
  <form method="POST" action="rmbzjin.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#800000"><br>
    </font>
  </div>
</form>
<font color="#800000">
<%else%>
</font><font color="#FFFFFF">
<br><br>
<%end if%>
<%if taiqiudian>0 then%>     
</font>    
  <form method="POST" action="jinzrmb5.asp" name="aff" onsubmit="bbb.disabled=1">    
  <div align="center"><font color="#FFFFFF"><img border="0" src="../chat/img/251.GIF" width="32" height="32">金币：<%=taiqiudian%>个，金币转换法力<%=int(taiqiudian*200)%>点<br>   
    <input type="text" name="input1" size="10" maxlength="5">   
    <input type="submit" value="转换" name="bbb" class="p9" >     
    </font>   
  </div>   
</form>   
<font color="#FFFFFF">   
<%else%>   
金币为0不能转换<br><br>   
<%end if%>   
 <b>法力转换</b>     
</font>   
                </div>  
            </td>  
          </tr>  
          <tr>   
            <td valign="top" bgcolor="#800000" >   
              <p><font color="#FFFF00">在这里你可以将金币，换成江湖法力！<br>转换：1个金币/5000点武功!</font></p> 
            </td> 
          </tr> 
        </table> 
<div align="center">
</div> 
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body> 
</html>