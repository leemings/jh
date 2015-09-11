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
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 银币,充值卡 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
hyrmb=rs("银币")
taiqiudian=rs("充值卡")
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
<html>

<head>
<link rel="stylesheet" href="../../css.css">
<title>银币充值</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="bg.gif">
<div align="center"><font color="#000000"><b>欢迎<font color=blue><%=aqjh_name%></font>光临银币充值！</b></font></div>
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
  <form method="POST" action="czkyb.asp" name="aff" onsubmit="bbb.disabled=1">     
  <div align="center">充值卡：<%=taiqiudian%>张，充值卡充值银币<%=int(taiqiudian*5000)%>个<br>    
    <input type="text" name="input1" size="10" maxlength="5">    
    <input type="submit" value="转换" name="bbb" class="p9" >    
  </div>    
</form>    
<%else%>    
充值卡为0不能充值<br><br>    
<%end if%>    
 <b><font color="#000000">银币充值</font></b>    
                </div>   
            </td>   
          </tr>   
          <tr>    
            <td valign="top" >    
              <p>在这里你可以将充值卡，充值金币！<br>转换：1张充值卡/5000个银币!</p> 
            </td> 
          </tr> 
        </table> 
<div align="center"><b><font color="#0000FF"><font color="#00FF66"><%=Application("aqjh_chatroomname")%></font></font></b></div> 
</body> 
</html>