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
rs.open "select 会员金卡,法力,轻功 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
hyrmb=rs("会员金卡")
qg=rs("轻功")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<link rel="stylesheet" href="../../css.css">
<title>轻功转换</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#000000">
<div align="center"><b><font color="#FF0000">&nbsp;欢迎<font color="blue"><%=aqjh_name%></font>光临轻功转换处！&nbsp;</font></b></div>
<div align="center"><b><a href="index.asp"><font color="#FF0000" size="2">返回转换市场</font></a></b></div>
<br>
        <table width="380" align="center" bordercolor="#339966" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" width="370" > 
              <font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <img border="0" src="jb.gif"> 
              ====》<img border="0" src="qg.gif"></font>
            </td>
          </tr>
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" width="370" > 
              <div align="center">
<%if hyrmb>1 then%>
  <form method="POST" action="jhq.asp" name="af" onsubmit="bb.disabled=1">
  <div align="center"><font color="#FF0000" size="2">金卡：<%=hyrmb%>元，金卡转换轻功<%=hyrmb*5000%>点<br>
    <input type="text" name="input" size="10" maxlength="10">
    <input type="submit" value="转换" name="bb" class="p9">     
    </font>    
  </div>  
</form>  
<font color="#FF0000" size="2">  
<%else%>  
会员金卡为1元不能转换<br><br> 
<%end if%> 
 轻功转换</font>    
                </div> 
            </td> 
          </tr> 
          <tr>  
            <td valign="top" bgcolor="#000000" width="370" >  
              <p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
              金卡转换轻功比例为1：50000点</font></p> 
            </td> 
          </tr> 
        </table> 
<p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<img border="0" src="gem01.gif"></font><font size="2" color="#0000FF">爱情江湖独创 
2005-01-12</font></p>
</body> 
</html>