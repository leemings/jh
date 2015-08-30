<html>
<head>
<title>快乐江湖管理♀wWw.happyjh.com♀</title>
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
<font style="FONT-SIZE: 9pt" color="#00ffff"> 【列出到<%=now()%>为止的非官府,掌门在线前30排名】<br>
--------------------------------------------------------------------------------<br>
［注意事项］<br>
<br>
1. 此处列出的开始计算日期为用户注册日期。<br> 
2. 如果你在排行榜的前三名那就有机会申请加入官府，如果你在排行的四、五名，那就有机会申请建立门派。<br> 
3. 申请加入官府必须在有官府被降职后，才能申请。<br> 
4. 申请建立门派必须有门派倒闭后，才可以申请。<br> 
<font style="FONT-SIZE: 9pt" color="#00ffff"><font color="#FF0000"><b>注意：申请每一个项目必须在每星期（掌门和官府）降职结果出来后，才可以申请，其它时间均为提醒，不作奖励</b></font></font></font><font style="FONT-SIZE: 9pt" color="#00ffff"><br>
--------------------------------------------------------------------------------<br>
<a href="TOP47.ASP">[列官府总在线排名]</a> <a href="TOP48.ASP">[列掌门总在线排名]</a> 
<a href="TOP49.ASP">[列出非官府,非掌门在线前30排名]</a> <a href="TOP50.ASP">[按平均每天在线列出所有用户后30排名]</a> </font>  
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
conn.open Application("sjjh_usermdb")
const MaxPerPage=10 
sql="select * from 用户 where 门派<>'官府' and 身份<>'掌门' order by mvalue desc"       
Set rs= Server.CreateObject("ADODB.Recordset")        
rs.open sql,conn,1,1        
 if rs.eof and rs.bof then        
       response.write "<p align='center'>没有可排行的对象 </p>"        
   else        
%>
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
      <font color="#0000FF">申请聊管</font> 
      <%elseif filename<=5 then%>
      <font color="#FF0000">申请门派</font> 
      <%else%>
      继续努力 
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