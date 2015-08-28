<%@ LANGUAGE=VBScript%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
%>

<html>
<head>
<link rel=stylesheet href='../../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../../chatroom/bg1.gif' bgproperties="fixed" leftmargin="3" topmargin="0">
<div align="center"><br>
<font size="3" color="#FFFFFF"><br>
  </font><b>要塞武器</b>
  <p><br>   
<!--#include file="data.asp"--><%  
sql="select 等级,攻击,防御,名称,状态 from mj where 所有者='"&username&"' and 类型='轩' order by 等级"
set rs=connb.execute(sql) 
if not(rs.EOF or rs.BOF) then    
%>   
  </p>
  <table cellpadding="3" cellspacing="1" border="0" align="center" width="100%">   
<%   
do while not (rs.eof or rs.bof)
if rs("状态")=true then lx="攻击"
if rs("状态")=false then lx="防御"
%>    
    <tr align="center">    
      <td colspan="2" height="20">武器：<a href="javascript:parent.talkfrm.settalk('//攻打','<%=rs("名称")%>')" title=等级：<%=rs("等级")%>><%=rs("名称")%></a> <br>           
      </td>           
    </tr>                   
        <tr align="center">    
      <td colspan="2" height="20"><%=lx%>：<%=rs(""&lx&"")%><br>           
      </td>           
    </tr> 
    <tr>            
      <td colspan="2">            
<hr noshade size="1" color=darkred>         
      </td>           
    </tr>           
               
    <%                 
rs.movenext           
loop           
rs.Close           
set rs=nothing           
connb.close           
set connb=nothing           
%>            
  </table>           
<table width="100%" border="0" cellspacing="2" cellpadding="2">           
<tr>           
      <td>这是专门用来攻打要塞的武器,想购置新的武器,请去<a href="#" onClick="window.open('../ys/add.asp','reg','width=580,height=290')"><font color="#F71004">要塞武器店</font></a></td>           
</tr>           
</table>           
<br>           
  <%           
else           
%>           
Sorry！您现在没有攻打要塞的武器，去设置一个吧！<a href="#" onClick="window.open('../ys/add.asp','reg','width=580,height=290')"><font color="#F71004">要塞武器店</font></a>               
  <%end if%>                
</div>                
</body>                
</html>          
