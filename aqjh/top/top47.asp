<html>
<head>
<title>��������</title>
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
<font style="FONT-SIZE: 9pt" color="#00ffff"> ���г���<%=now()%>Ϊֹ�Ĺٸ�������������<br>
--------------------------------------------------------------------------------<br>
�۹ٸ�ע�������<br>
<br>
1. �˴��г����г�Ϊ�ٸ����վ����������վ���������Щ���ݰ��Źٸ��ĸ�λ��<br>  
2. �ٸ��������Ĺ������ţ�ֻ������ʱ�䳤��ʱʱ�̹̿����������Ż�ʹ�����������ȶ�����ÿһλ��Ҷ����ܵ����֡�<br>  
3. �ٸ�Ҫ�����������ݸ����Լ���ȱ�㣬��֤�����վ��������ÿһ������<br>  
4. �ٸ�ÿ���������5��Сʱ��<br>  
5. �ٸ�����ÿ�������µ���һ�Σ��Ѳ����������Ĺٸ���ְ��������ιٸ�ʤ�ο�ȱ��λ��<br>  
</font><b>
<font style="FONT-SIZE: 9pt" color="#FF6600"> ע�⣺�ٸ����ڹ涨��ʱ�����µ�����������ٸ������������������Ҳ����ʱ������</font></b><br>
<font style="FONT-SIZE: 9pt" color="#00ffff">--------------------------------------------------------------------------------------------------------------------------<br>
<a href="TOP47.ASP">[</font><font style="FONT-SIZE: 9pt" color="#FFFF00">�йٸ�����������</font><font style="FONT-SIZE: 9pt" color="#00ffff">]  
  </font> 
</a> <font style="FONT-SIZE: 9pt" color="#00ffff"> <a href="TOP48.ASP">[</a></font><a href="TOP48.ASP"><font style="FONT-SIZE: 9pt" color="#FF0000">��������������]   
  </font> 
</a><font style="FONT-SIZE: 9pt" color="#00ffff">  
<a href="TOP49.ASP">[�г��ǹٸ�,����������ǰ30����]</a> <a href="TOP50.ASP">[</font><font style="FONT-SIZE: 9pt" color="#FF6600">��ƽ��ÿ�������г������û���30����</font><font style="FONT-SIZE: 9pt" color="#00ffff">]   
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
conn.open Application("aqjh_usermdb")
const MaxPerPage=10
sql="select * from �û� where ����='�ٸ�' and ����<>'����' order by mvalue desc"   
Set rs= Server.CreateObject("ADODB.Recordset")        
rs.open sql,conn,1,1        
 if rs.eof and rs.bof then        
       response.write "<p align='center'>û�п����еĶ��� </p>"        
   else        
%>

<TABLE width='99%' ALIGN=center CELLSPACING=2 CELLPADDING=5>        
  <tr><td>        
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>        
   <tr bgcolor="#C4DEFF">        
    <td align="center" bgcolor="#000000"><font color="#FFFFFF">����<img border="0" src="../chat/img/enemyin05.gif"></font></td>        
    <td align="center">����</td>        
    <td align="center">����<img border="0" src="../chat/img/02.jpg"></td>        
    <td align="center">����</td>        
    <td align="center">ע��ʱ��</td>    
    <td align="center">����¼</td>    
    <td align="center">�»���</td>       
    <td align="center">����</td>       
    <td align="center">�վ�����</td>       
    <td align="center" bgcolor="#FFFFFF">��ʾ<img border="0" src="../chat/img/hpsy.gif"></td>       
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
    <td align="center" bgcolor="#FFFFFF"><font color="#000000"><%=rs("����")%></font></td>       
    <td align="center"><%=rs("����")%></td>       
    <td align="center"><%=rs("����")%></td>       
    <td align="center"><%=rs("grade")%></td>    
    <td align="center"><%=rs("regtime")%></td>      
    <td align="center"><%=rs("lasttime")%></td>      
    <td align="center"><font color=red><%=rs("mvalue")%></font></td>           
    <td align="center"><font color=red><%=ts%></font></td>           
    <td align="center"><font color=red><%=zx%>����</font></td>           
    <td align="center" bgcolor="#FFFFFF"> 
      <%if zx>=300 then%>
      ��������  
      <%elseif zx<300 then%>
      <font color="#FF0000">������ҡ</font>  
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
<p align="center"><br><a href="javascript:self.close()"><font style="FONT-SIZE: 9pt">���رմ��ڡ�</font></a>      
</body>      
</html>