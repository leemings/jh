<%@ LANGUAGE=VBScript codepage ="936" %>
<script lanaguage=javascript>
if(window.name!="aqjh_win"){ var i=1;while (i<=50){window.alert("你想作什么呀，你倒是刷啊！哈！慢慢点50次！！");i=i+1;}top.location.href="../../exit.asp"}
</script>
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
rs.open "select 银两,金币,等级 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
hyrmb=rs("银两")
taiqiudian=rs("金币")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<link rel="stylesheet" href="../../css.css">
<title>银两转换</title>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#000000">
<div align="center"><b><font color="#FF0000">欢迎大侠<font color=blue><%=session("aqjh_name")%></font>光临银两转换处！</font></b></div>
<div align="center"><b><a href="index.asp"><font color=#FF0000 size="2">返回转换市场</font></a></b></div>
<br>
        <table width="300" align="center" bordercolor="#FF0000" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" > 
              <div align="center">
<%if hyrmb>0 then%>
<p><font size="2"><br>
<%end if%>
 <img border="0" src="jb.gif"> ==&gt;&gt; <img border="0" src="yl.gif"><%if taiqiudian>0 then%></font></p> 
  <form method="POST" action="yinzhok.asp" name="aff" onsubmit="bbb.disabled=1">    
  <div align="center"><font color="#FF0000" size="2">金币：<%=taiqiudian%>个，金币转换银两<%=int(taiqiudian*1000000)%>两    
    </font>   
    <font color="#FFFFFF" size="2"><br>   
    <input type="text" name="input1" size="10" maxlength="5">   
    <input type="submit" value="转换" name="bbb" class="p9" >     
    </font>    
  </div>    
</form>    
<font color="#FFFFFF" size="2">    
<%else%>    
</font><font color="#0000FF" size="2">    
金币为0不能转换     
</font>    
<font color="#FFFFFF" size="2">    
<br><br>    
<%end if%>    
</font>    
<font color="#FF0000" size="2">    
银<b>两转换</b>     
</font>    
                </div>   
            </td>   
          </tr>   
          <tr>    
            <td valign="top" bgcolor="#000000" >    
              <p><font color="#FFFF00" size="2">在这里你可以将金币，换成江湖银两！<br>转换：1个金币/1000000银两!</font></p>   
            </td>   
          </tr>   
        </table>   
<p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
<img border="0" src="gem01.gif"></font><font size="2" color="#0000FF">爱情江湖独创  
2005-01-12</font></p> 
</body>   
</html>