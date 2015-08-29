<html>
<head>
<title>e线江湖管理♀wWw.51eline.com♀</title>
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
<font style="FONT-SIZE: 9pt" color="#00ffff"> 【列出到<%=now()%>为止的官府总在线排名】<br>
--------------------------------------------------------------------------------<br>
［官府注意事项］<br>
<br>
1. 此处列出所有成为官府的日均在线情况，站长会根据这些数据安排官府的岗位。<br>  
2. 官府是社区的管理部门，只有在线时间长，时时刻刻管理社区，才会使社区的秩序稳定，让每一位玩家都感受到快乐。<br>  
3. 官府要根据所列数据改正自己的缺点，保证能完成站长交给的每一项任务。<br>  
4. 官府每天必须在线5个小时。<br>  
5. 官府会在每星期重新调整一次，把不符合条件的官府撤职，重新提拔官府胜任空缺岗位。<br>  
</font><b>
<font style="FONT-SIZE: 9pt" color="#FF6600"> 注意：官府会在规定的时间重新调整，但如果官府不能作到让玩家满意也会随时开除！</font></b><br>
<font style="FONT-SIZE: 9pt" color="#00ffff">--------------------------------------------------------------------------------------------------------------------------<br>
<a href="TOP47.ASP">[</font><font style="FONT-SIZE: 9pt" color="#FFFF00">列官府总在线排名</font><font style="FONT-SIZE: 9pt" color="#00ffff">]  
  </font> 
</a> <font style="FONT-SIZE: 9pt" color="#00ffff"> <a href="TOP48.ASP">[</a></font><a href="TOP48.ASP"><font style="FONT-SIZE: 9pt" color="#FF0000">掌门总在线排名]   
  </font> 
</a><font style="FONT-SIZE: 9pt" color="#00ffff">  
<a href="TOP49.ASP">[列出非官府,非掌门在线前30排名]</a> <a href="TOP50.ASP">[</font><font style="FONT-SIZE: 9pt" color="#FF6600">按平均每天在线列出所有用户后30排名</font><font style="FONT-SIZE: 9pt" color="#00ffff">]   
  </font> 
</a>  
<font style="FONT-SIZE: 9pt" color="#00ffff">   
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
conn.open Application("sjjh_usermdb")
const MaxPerPage=10
sql="select * from 用户 where 门派='官府' and 身份<>'掌门' order by mvalue desc"   
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
    <td align="center" bgcolor="#000000"><font color="#FFFFFF">姓名<img border="0" src="../chat/img/enemyin05.gif"></font></td>        
    <td align="center">门派</td>        
    <td align="center">身份<img border="0" src="../chat/img/02.jpg"></td>        
    <td align="center">级别</td>        
    <td align="center">注册时间</td>    
    <td align="center">最后登录</td>    
    <td align="center">月积分</td>       
    <td align="center">天数</td>       
    <td align="center">日均在线</td>       
    <td align="center" bgcolor="#FFFFFF">提示<img border="0" src="../chat/img/hpsy.gif"></td>       
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
    <td align="center" bgcolor="#FFFFFF"><font color="#000000"><%=rs("姓名")%></font></td>       
    <td align="center"><%=rs("门派")%></td>       
    <td align="center"><%=rs("身份")%></td>       
    <td align="center"><%=rs("grade")%></td>    
    <td align="center"><%=rs("regtime")%></td>      
    <td align="center"><%=rs("lasttime")%></td>      
    <td align="center"><font color=red><%=rs("mvalue")%></font></td>           
    <td align="center"><font color=red><%=ts%></font></td>           
    <td align="center"><font color=red><%=zx%>分钟</font></td>           
    <td align="center" bgcolor="#FFFFFF"> 
      <%if zx>=300 then%>
      继续保持  
      <%elseif zx<300 then%>
      <font color="#FF0000">宝座动摇</font>  
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
<p align="center"><br><a href="javascript:self.close()"><font style="FONT-SIZE: 9pt">『关闭窗口』</font></a>      
</body>      
</html>