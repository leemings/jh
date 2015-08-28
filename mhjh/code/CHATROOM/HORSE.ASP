<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("yx8_mhjh_connstr")
rs.open "select * FROM d WHERE c='马'",conn
xzmd=""
do while not rs.bof and not rs.eof
 xzmd=xzmd&"<option >"&rs("xm")&" "& rs("ydx") &"号"& rs("yl") &"</option>"
rs.movenext
loop
tmprs=conn.execute("Select count(*) As 数量 from d where c='马'")
dars=tmprs("数量")
rs.close
rs.open "select 会员 FROM 用户 WHERE 姓名='"&username&"'",conn
hy=rs("会员")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<head>
<link rel="stylesheet" href="css.CSS" type="text/css">
<title>在线跑马</title>
</head>
<body background='bg1.gif' oncontextmenu=self.event.returnValue=false >
<form method="post" action="horseok.asp">
  <table width="102%" border="0" cellspacing="2" cellpadding="2" align="center">
    <tr align="center"> 
      <td><font face="隶书" size="4">在线赌马</font></td>
</tr>

<tr align="center"> 
      <td height="3">
        <select name="1111111">
          <option selected>&nbsp;&nbsp;&nbsp;&nbsp;下&nbsp;注&nbsp;名&nbsp;单&nbsp;&nbsp;</option>
         <%=xzmd%>
        </select>
        <br>
        现在还有<%=(10-dars)%>个开马!</td>
</tr>
<tr align="center"> 
<td> 
<select name="money">
<option value="10000" selected>1万</option>
<option value="100000">10万</option>
<option value="300000">30万</option>
<option value="500000">50万</option>
<% if hy=true then %>
<option value="50000">20万</option>
<option value="500000">50万</option>
<option value="1000000">100万</option>
<%end if%>
</select> 
</td> 
</tr> 
<tr align="center">  
<td>选择号码</td> 
</tr> 
<tr align="center">  
<td>  
<font size="3" color="#FF0000">  
[1]</font>  <input type="radio" name="radiobutton" value="1" checked> 
<img src="horse/1.GIF" width="36" height="35"> <br>   
<font size="3" color="#FF0000">  
[2]</font>  <input type="radio" name="radiobutton" value="2">  
<img src="horse/2.GIF" width="36" height="35"><br>   
<font size="3" color="#FF0000">  
[3]</font>  <input type="radio" name="radiobutton" value="3">  
<img src="horse/3.GIF" width="36" height="35"><br>   
<font size="3" color="#FF0000">  
[4]</font>  <input type="radio" name="radiobutton" value="4">  
<img src="horse/4.GIF" width="36" height="35"><br>   
<font size="3" color="#FF0000">  
[5]</font>  <input type="radio" name="radiobutton" value="5">  
<img src="horse/5.GIF" width="36" height="35"><br>   
      </td>   
</tr>   
<tr align="center">    
<td>    
<input type="submit" name="Submit" value="确定">   
<input type="button" name="Submit2" value="返回" onClick=this.document.location.href='onlinelist.asp'>   
</td>   
</tr>   
<tr align="center">    
<td>最小投注额为10000<br>   
        非会员最大投注额为<font color="red">50</font>万<br>   
        会员最大投注额为<font color="red">100</font>万<br>   
        下注够10人自动开奖</td>   
</tr>   
</table>   
</form>