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
rs.open "select 银两,金币 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("银两")
taiqiudian=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>银两转换♀wWw.happyjh.com♀</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../chat/back.jpg">
<div align="center"><b><font color="#008080">欢迎大侠<img border="0" src="../chat/img/mob_s_20.gif">光临银两转换处</font><font color="#000000">！</font></b></div>
<br>
        <table width="300" align="center" bordercolor="#FF0000" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#008080" > 
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
  <form method="POST" action="jinzrmb2.asp" name="aff" onsubmit="bbb.disabled=1">   
  <div align="center"><font color="#FFFFFF">金币：<%=taiqiudian%>个，金币转换银两<%=int(taiqiudian*1000000)%>两<br>  
    <input type="text" name="input1" size="10" maxlength="5">  
    <input type="submit" value="转换" name="bbb" class="p9" >   
    </font>  
  </div>  
</form>  
<font color="#FFFFFF">  
<%else%>  
金币为0不能转换<br><br>  
<%end if%>  
 <img border="0" src="../chat/img/251.GIF">银<b>两转换</b>   
</font>  
                </div> 
            </td> 
          </tr> 
          <tr>  
            <td valign="top" bgcolor="#008080" >  
              <p><font color="#FFFF00">在这里你可以将金币，换成江湖银两！<br>转换：1个金币/1000000银两!</font></p> 
            </td> 
          </tr> 
        </table> 
<div align="center"> 
</div> 
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body> 
</html>