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
<p><font style="FONT-SIZE: 9pt" color="#00ffff"> ���г���<%=now()%>Ϊֹ�������ʺź�30������<br>
  --------------------------------------------------------------------------------<br>
  ��ע�������<br>
  <br>
  <font color="#FF0000" size="4"><b>�ڴ��г����ǲ��������ߵ��ʺţ�վ���ᶨʱɾ��һ�����ʺţ��ڴ˵�ǰ�����ʺ�Ҫע��~��</b></font></font><font style="FONT-SIZE: 9pt" color="#00ffff"><br>
  --------------------------------------------------------------------------------<br>
<a href="TOP47.ASP">[�йٸ�����������]</a> <a href="TOP48.ASP">[����������������]</a> 
<a href="TOP49.ASP">[�г��ǹٸ�,����������ǰ30����]</a> <a href="TOP50.ASP">[��ƽ��ÿ�������г������û���30����]</a>  
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
sql="select * from �û� where ����<>'�ٸ�' and ����<>'����' order by -mvalue desc"       
Set rs= Server.CreateObject("ADODB.Recordset")        
rs.open sql,conn,1,1        
 if rs.eof and rs.bof then        
       response.write "<p align='center'>û�п����еĶ��� </p>"        
   else        
%>
</p>
<TABLE width='99%' ALIGN=center CELLSPACING=2 CELLPADDING=5>        
  <tr><td>        
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>        
   <tr bgcolor="#C4DEFF">        
    <td align="center">����</td>        
    <td align="center">����</td>        
    <td align="center">����</td>        
    <td align="center">����</td>        
    <td align="center">ע��ʱ��</td>    
    <td align="center">����¼</td>    
    <td align="center">�ۻ���</td>       
    <td align="center">����</td>       
    <td align="center">�վ�����</td>       
    <td align="center">��ʾ</td>       
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
    <td align="center"><%=rs("����")%></td>       
    <td align="center"><%=rs("����")%></td>       
    <td align="center"><%=rs("����")%></td>       
    <td align="center"><%=rs("grade")%></td>    
    <td align="center"><%=rs("regtime")%></td>      
    <td align="center"><%=rs("lasttime")%></td>      
    <td align="center"><font color=red><%=rs("mvalue")%></font></td>           
    <td align="center"><font color=red><%=ts%></font></td>           
    <td align="center"><font color=red><%=zx%>����</font></td>           
    <td align="center"> 
      <%if filename<=3 then%>
      <font color="#0000FF">ע��ע��</font> 
      <%elseif filename<=5 then%>
      <font color="#FF0000">Σ��Σ��</font> 
      <%else%>
      �Ƚ����� 
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
<p align="center"><br><align="center"><a href="javascript:self.close()"><font style="FONT-SIZE: 9pt">���رմ��ڡ�</font></a>      
</body>      
</html>