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
rs.open "select 会员金卡,法力 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
hyrmb=rs("会员金卡")
fl=rs("法力")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>法力转换♀wWw.happyjh.com♀</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../BG.GIF">
<div align="center"><b><font color="#008080">欢迎<%=sjjh_name%><img border="0" src="../chat/img/tu2.gif">光临法力转换处！</font></b></div>
<br>
        <table width="300" align="center" bordercolor="#FF0000" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#339966" > 
              <div align="center">
<%if hyrmb>1 then%>
  <form method="POST" action="jhf.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#FFFFFF">金卡：<%=hyrmb%>元，金卡转换法力<%=hyrmb*50000%>点<br>
    <input type="text" name="input" size="10" maxlength="10">
    <input type="submit" value="转换" name="bb" class="p9">  
    </font> 
  </div>
</form>
<font color="#FFFFFF">
<%else%>
会员金卡为1元不能转换<br><br>
<%end if%>
<%if fl>100000 then%>   
</font>  
  <form method="POST" action="fhj.asp" name="aff" onsubmit="bbb.disabled=1"> 
  <div align="center"><font color="#FFFFFF">法力：<%=fl%>点，法力转换金卡<%=int(fl/100000)%>元<br>
    <input type="text" name="input1" size="10" maxlength="10">
    <input type="submit" value="转换" name="bbb" class="p9" >  
    </font> 
  </div>
</form>
<font color="#FFFFFF">
<%else%>
法力为100000点不能转换<br><br>
<%end if%>
 法力转换</font>  
                </div>
            </td>
          </tr>
          <tr> 
            <td valign="top" bgcolor="#339966" > 
              <p><font color="#FFFF00">在这里你可以将金卡转换成法力，也可以将法力转换成金卡！<br>转换：1元/50000点法力，100000点法力/1元!同样您也可以将法力转换成现金购买卡片！</font></p>
            </td>
          </tr>
        </table>
<div align="center"><b><font color="#FF0000"><%=Application("sjjh_chatroomname")%></font></b>
</div>
</body>
</html>