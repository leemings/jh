<%
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if Application("yx8_mhjh_biwu")<>"" then %>
<HTML>
<HEAD>
<style>
img{border:0}
input,td{font-size:9pt}
p{color:white}
.nb{border-left:1px solid #000000;border-right:1px solid #ffffff;border-top:1px solid #000000;border-bottom:1px solid #ffffff;font-size: 9pt;background:lavender}
</style>
<script language=javascript>setTimeout('this.location.reload();',60000);</script>
<title></title>
</HEAD>

<body background="../bg1.gif" >
<CENTER>
<p></p>
<br>
<br>
<table>
<tr>
<td width=150>
<b>魔导论武场</b>
<tr>
<td width=150>
<% 
if Application("yx8_mhjh_hong")<>"无" then 
   set conn=server.createobject("adodb.connection")  
   conn.open Application("yx8_mhjh_connstr") 
   set rst=server.CreateObject ("adodb.recordset") 
   sql="select 攻击,防御,体力,内力 from 用户 where 姓名='"&Application("yx8_mhjh_hong")&"'" 
   set rs=conn.execute(sql) 
   gj=rs("攻击") 
   fy=rs("防御") 
   tl=rs("体力") 
   nl=rs("内力") 
   pla=int(10000000/gj+10000000/fy+10000000/tl+5000000/nl) 
   wer=""&pla&"" 
else
   wer="1" 
end if 
if Application("yx8_mhjh_hei")<>"无" then   
   set conn=server.createobject("adodb.connection")    
   conn.open Application("yx8_mhjh_connstr")   
   set rst=server.CreateObject ("adodb.recordset")   
   sql="select 攻击,防御,体力,内力 from 用户 where 姓名='"&Application("yx8_mhjh_hei")&"'"   
   set rs=conn.execute(sql)   
   gj=rs("攻击")   
   fy=rs("防御")   
   tl=rs("体力")   
   nl=rs("内力")   
   plb=int(10000000/gj+5000000/fy+10000000/tl+5000000/nl)   
   wfr=""&plb&""      
else  
   wfr="1"  
end if   
if wer<>"1" and wfr<>"1" and wer>wfr then
   msg1="3"
   msg2="2"
elseif wer<>"1" and wfr<>"1" and wer<wfr then
   msg1="2"
   msg2="3"
else
   msg1="1"
   msg2="1"

end if
%>  
场次：<input type=text size=12 value=<%=Application("yx8_mhjh_changci")%> class=nb name=n3 readonly><br>
红方：<input type=text size=12 value=<%=Application("yx8_mhjh_hong")%> class=nb name=n1 readonly><br>
赔率：<input type=text value=1：<%=msg1%> name=pa size=12 class=nb readonly><br>  
黑方：<input type=text size=12 value=<%=Application("yx8_mhjh_hei")%> class=nb name=n2 readonly><br>  
赔率：<input type=text value=1：<%=msg2%> name=pb size=12 class=nb readonly><br><br>   
<tr>   
<td width=150>   
<form name=j action=Join.asp  method=post>   
<input type=submit name=ac1 value=参加比武><br>   
<input type=submit name=ac2 value=我要认输><br><br>  
<select size=1 name=xz class=nb><option value="<%=Application("yx8_mhjh_hong")%>" selected>我买红赢<option value="<%=Application("yx8_mhjh_hei")%>">我买黑赢</select><br>  
<select size=1 name=zs class=nb><option value=100 selected>100两</option><option value=300>300两</option><option value=500>500两</option>
<%
   set conn=server.createobject("adodb.connection")  
   conn.open Application("yx8_mhjh_connstr") 
   set rst=server.CreateObject ("adodb.recordset") 
   sql="select 会员 from 用户 where 姓名='"&username&"'" 
   set rs=conn.execute(sql) 
   hui=rs("会员")
   if hui=True then
%>
<option value=1000>1000两</option><option value=5000>5000两</option><option value=10000>10000两</option>
<% end if %>
</select>   
<input type=hidden name=re value="<%=msg1%>">  
<input type=hidden name=rg value="<%=msg2%>">  
<input type=submit name=ac3 value=我要下注><br>   
</form>   
<tr>    
<td width=150> 
<input type=button value='我要领奖' onclick=javascript:location.href='dui.asp';><br><br>   
非会员最大下注额为500两,会员为10000两  
</table>  
</CENTER>  
</body>  
</html>
<% 
rs.Close                                    
set rs=nothing                                    
conn.Close                                    
set conn=nothing  
else
Response.write "<body background='../bg1.gif'><br><br><br><p align=center>比武场还没有开门<br><a href=../../chatroom/onlinelist.asp>返回</a></p></body> "
end if %>