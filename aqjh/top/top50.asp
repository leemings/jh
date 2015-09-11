<html>
<head>
<title>江湖管理</title>
<style>
p{font-size:9pt; color:#ffee00}
td, select, input{font-size:9pt; color:#000000; height:9pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=#3a4b91>
<p><font style="FONT-SIZE: 9pt" color="#00ffff"> 【列出到<%=now()%>为止的所有帐号后30排名】<br>
  --------------------------------------------------------------------------------<br>
  ［注意事项］<br>
  <br>
  <font color="#FF0000" size="4"><b>在此列出的是不经常在线的帐号，站长会定时删除一部分帐号，在此的前五名帐号要注意~！</b></font></font><font style="FONT-SIZE: 9pt" color="#00ffff"><br>
  --------------------------------------------------------------------------------<br>
<a href="TOP47.ASP">[列官府总在线排名]</a> <a href="TOP48.ASP">[列掌门总在线排名]</a> 
<a href="TOP49.ASP">[列出非官府,非掌门在线前30排名]</a> <a href="TOP50.ASP">[按平均每天在线列出所有用户后30排名]</a>  
  </font> 
  <%  
n=Year(date())  
y=Month(date())  
r=Day(date())  
s=Hour(time())  
f=Minute(time())  
m=Second(time())  
if len(y)=1 then y="0" & y  
if len(r)=1 then r="0" & r  
if len(s)=1 then s="0" & s  
if len(f)=1 then f="0" & f  
if len(m)=1 then m="0" & m  
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m        
dim sql        
dim rs        
dim filename    
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
const MaxPerPage=10
sql="select * from 用户 where 门派<>'官府' and 身份<>'掌门' order by -mvalue desc"       
Set rs= Server.CreateObject("ADODB.Recordset")        
rs.open sql,conn,1,1        
 if rs.eof and rs.bof then        
       response.write "<p align='center'>没有可排行的对象 </p>"        
   else        
%>
</p>
<TABLE width='99%' ALIGN=center CELLSPACING=2 CELLPADDING=5>        
  <tr><td>        
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>        
   <tr bgcolor="#C4DEFF">        
    <td align="center">姓名</td>        
    <td align="center">门派</td>        
    <td align="center">身份</td>        
    <td align="center">级别</td>        
    <td align="center">注册时间</td>    
    <td align="center">最后登录</td>    
    <td align="center">累积分</td>       
    <td align="center">天数</td>       
    <td align="center">日均在线</td>       
    <td align="center">提示</td>       
  </tr>       
<%       
do while not rs.eof and not rs.bof       
ts=DateDiff("d",rs("regtime"),rs("lasttime"))    
zf=rs("mvalue")    
if ts=0 then    
ts=1    
end if    
zx=int(zf/ts)
 filename=filename+1      
%>       
  <tr bgcolor=#f7f7f7>       
    <td align="center"><%=rs("姓名")%></td>       
    <td align="center"><%=rs("门派")%></td>       
    <td align="center"><%=rs("身份")%></td>       
    <td align="center"><%=rs("grade")%></td>    
    <td align="center"><%=rs("regtime")%></td>      
    <td align="center"><%=rs("lasttime")%></td>      
    <td align="center"><font color=red><%=rs("mvalue")%></font></td>           
    <td align="center"><font color=red><%=ts%></font></td>           
    <td align="center"><font color=red><%=zx%>分钟</font></td>           
    <td align="center"> 
      <%if filename<=3 then%>
      <font color="#0000FF">注意注意</font> 
      <%elseif filename<=5 then%>
      <font color="#FF0000">危险危险</font> 
      <%else%>
      比较幸运 
      <%end if%>
    </td>      
  </tr>       
<%       
 rs.movenext       
     
 if filename>29 then Exit Do       
 loop       
end if       
conn.close       
%>       
</table>       
</td></tr>       
</TABLE>   
<p align="center"><br><align="center"><a href="javascript:self.close()"><font style="FONT-SIZE: 9pt">『关闭窗口』</font></a>      
</body>      
</html>