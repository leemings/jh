<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=Request.QueryString("un")
mygrade=session("yx8_mhjh_usergrade")
if mygrade<5 then Response.Redirect "../error.asp?id=242"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from 用户 where 姓名='"&un&"'",conn
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=221"
%>
<html>
<head>
<link rel='stylesheet' href='css.css'>
<title>查看资料</title>
<Script Language="JScript" Runat=Server>
function autolevel(lvl){var i=0;temp="";for(i=0;i<=lvl;i++){temp=temp+"<img src=\"image\/star.gif\">";}return(temp);};
</Script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFDDF2" background='bg1.gif'>
<div align="center">
<center>
<table border="1" width="350" cellspacing="0" cellpadding="0">
<tr>
<td width="5%" rowspan="7" height="128">
<p align="center"><img border='0' src='<%=rs("头像")%>'></td>
<td width="75%" height="16" colspan="2">
<p align="center">&nbsp;
<marquee style="color: #FF0000">笑问天下谁傲世，唯有剑侠我独尊</marquee>
</p>
</td>
</tr>
<tr>
<td width="27%" height="16">头衔:<%=rs("头衔")%></td>
<td width="48%" height="16">会员:<%=rs("会员")%></td>
</tr>
<tr>
<td width="27%" height="16">等级:<%=rs("等级")%></td>
<td width="48%" height="16">门派:<%=rs("门派")%></td>
</tr>
<tr>
<td width="27%" height="16">积分:<%=rs("积分")%></td>
<td width="48%" height="16">身份:<%=rs("身份")%></td>
</tr>
<tr>
<td width="27%" height="16">国家:<%=rs("国家")%></td>
<td width="48%" height="16">攻击:<%=rs("攻击")%></td>
</tr>
<tr>
<td width="27%" height="16">职位:<%=rs("职位")%></td>
<td width="48%" height="16">防御:<%=rs("防御")%></td>
</tr>
<tr>
<td width="75%" height="46" colspan="2" rowspan="6">个人签名:<br><br><%=rs("签名档")%></td>
</tr>
<tr>
<td width="36%" height="18">侠客:<%=rs("姓名")%></td>
</tr>
<tr>
<td width="36%" height="19">荣誉:<%
if rs("等级")="" then response.write("武林新手")
if rs("等级")<=3 then response.write("武林新手")
if rs("等级")<=6 and rs("等级")>3  then response.write("武林游侠")
if rs("等级")<=8 and rs("等级")>6  then response.write("武林斗侠")
if rs("等级")<=10and rs("等级")>8  then response.write("初级战将")
if rs("等级")<=11 and rs("等级")>10  then response.write("中级战将")
if rs("等级")<=14 and rs("等级")>11  then response.write("高级战将")
if rs("等级")<=17 and rs("等级")>14  then response.write("超级战将")
if rs("等级")>17 then response.write("终极战神")
%>
</td>
</tr>
<tr>
<td width="36%" height="19">配偶:<%=rs("配偶")%></td>
</tr>
<tr>
<td width="36%" height="19">现金:<%=rs("银两")%>两</td>
</tr>
<tr>
<td width="36%" height="19">存款:<%=rs("存款")%>两</td>
</tr>
<tr>
<td width="111%" height="1" colspan="3">
星级:<%=autolevel(rs("等级")/5)%>
</td>
</tr>
<tr>
<td width="111%" height="1" colspan="3">
<p align="center">&nbsp;&nbsp;<a href="../dongfang/MARRIAGE.ASP" target="_blank">恋爱</a>&nbsp;&nbsp;                            
<a href="../OTHER/MARRIAGE.ASP" target="_blank">求婚</a>&nbsp; &nbsp; 挑战&nbsp;&nbsp;&nbsp;                            
<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("oicq")%>" target="_blank">OICQ</a>&nbsp;&nbsp;                            
<a href="mailto:<%=rs("电子邮箱")%>">邮箱</a></td>                           
</tr>                           
</table>                           
</center>                           
</div>                                        
</body>                                      
</html>                           
<%                    
rs.Close                    
set rs=nothing      
conn.Close                    
set conn=nothing                     
%>                           
