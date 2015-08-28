<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 银两,资质,攻击,防御,道德 from 用户 where 姓名='"&username&"'"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>江湖募捐</title>
<link rel="stylesheet" href="../Style.css">
</head>
<body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false>
<p align="center"><font face="隶书" size="3"><b>江 湖  
募 捐</b></font></p>   
<p align="center">　</p>   
<form method="post" action="mj.asp">   
  <table width="400" border="1" cellspacing="0" cellpadding="5" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">   
    <tr>    
      <td>募捐人：</td>   
      <td><%=username%>&nbsp;</td>   
      <td>资质:<%=rs("资质")%></td>   
    </tr>   
    <tr>    
      <td>现在银两：</td>   
      <td><%=rs("银两")%>&nbsp;两</td>   
      <td>攻击:<%=rs("攻击")%></td>   
    </tr>   
    <tr>    
      <td>募捐金额:</td>   
      <td>    
        <input type="text" name="mj" size="10" maxlength="10">   
        最小金额1000</td>    
      <td>防御:<%=rs("防御")%></td>    
    </tr>    
    <tr>     
      <td colspan="2" >     
        <div align="center">     
          <input type="submit" name="Submit" value="募 捐">    
          　     
          <input type="reset" name="Submit2" value="清 空">    
        </div>    
      </td>    
      <td>道德:<%=rs("道德")%></td>    
    </tr>    
  </table>    
</form>    
<div align="center"><br>    
  <font size=2>募捐可对你有大大的好处哟，钱多了没地方花可以捐给社会大众嘛！<br>   
  捐1000两白银会加道德加20、资质会提升10哟<br>  
 </font><font color=#999999 size=2><br>  
  <br>  
</div>  
</body>  
</html>  
<% 
rs.Close              
set rs=nothing              
conn.Close              
set conn=nothing    
%>