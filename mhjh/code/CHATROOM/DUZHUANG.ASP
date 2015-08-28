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
rs.open "select * FROM d WHERE c='银' and sf<>'庄家'",conn
xzmd=""
do while not rs.bof and not rs.eof
 xzmd=xzmd&"<option >"&rs("xm")&" "& rs("ydx") &" "& rs("yl") &"</option>"
rs.movenext
loop
tmprs=conn.execute("Select count(*) As 数量 from d where c='银' and sf<>'庄家'")
dars=tmprs("数量")
rs.close
rs.open "select 会员 FROM 用户 WHERE 姓名='"&username&"'",conn
hy=rs("会员")
rs.close
%>

<head>
<link rel="stylesheet" href="css.CSS" type="text/css">
<title>在线赌银两</title>
</head>
<body background='bg1.gif' oncontextmenu=self.event.returnValue=false >
<form method="post" action="duzhuangok.asp">
  <table width="102%" border="0" cellspacing="2" cellpadding="2" align="center">
    <tr align="center"> 
      <td><font face="隶书" size="4">在线赌博</font></td>
</tr>

<tr align="center"> 
      <td height="3">
        <select name="1111111">
          <option selected>---下&nbsp;注&nbsp;名&nbsp;单---</option>
         <%=xzmd%>
        </select>
        <br><%rs.open "select * FROM d WHERE c='银' and sf='庄家'",conn
             if not(rs.eof) or not(rs.bof) then
              xm=rs("xm")%>
        庄家：<%=xm%>
             <%else%>
        现在没有庄家
             <%end if
             rs.close
             set rs=nothing
             conn.close
             set conn=nothing
             %>     
        <br>
        现在还有<%=(5-dars)%>个开赌局</td>
</tr>
<tr align="center"> 
<td> 
<input type=button value='赌场坐庄' onClick=javascript:location.href="zuozuang.asp" name="button2232"><br>
<input type=button value='赌场吆喝' onClick=javascript:location.href="yaohe.asp"  name="button2233"><br>
<select name="money">
<option value="100000" selected>10万</option>
<option value="300000">30万</option>
<option value="500000">50万</option>
<option value="2000000">200万</option>
<% if hy=true then %>
<option value="500000">50万</option>
<option value="2000000">200万</option>
<option value="5000000">500万</option>
<%end if%>
</select> 
</td> 
</tr> 
<tr align="center">  
<td>选择号码</td> 
</tr> 
<tr align="center">  
<td>  
&nbsp;  <input type="radio" name="id" value="大" checked>
 <img border="0" src="image/da.GIF"> <br>  
&nbsp;  <input type="radio" name="id" value="小"> 
 <img border="0" src="image/xiao.GIF"><br>  
&nbsp;  <input type="radio" name="id" value="单"> 
 <img border="0" src="image/dan.gif"><br>  
&nbsp;  <input type="radio" name="id" value="双"> 
 <img border="0" src="image/shuang.gif"><br>   
      </td>  
</tr>  
<tr align="center">   
<td>   
<input type="submit" name="Submit" value="确定"> 
<input type="button" name="Submit2" value="返回" onClick=this.document.location.href='onlinelist.asp'>    
</td>  
</tr>  
<tr align="center">   
<td>最小投注额为1000<br>  
        非会员最大投注额为<font color="red">200</font>万<br>  
        会员最大投注额为<font color="red">500</font>万<br>  
        下注够5人自动开奖<br><img border="0" src="../image/251.GIF"></td>  
</tr>  
</table>  
</form>
