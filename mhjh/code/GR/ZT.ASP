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
sql="Select  * from 用户 where 姓名='"&username&"'"
set rs=conn.Execute(sql)
%>
<head>
<link href="../style2.css" rel="stylesheet" type="text/css">
<title>个人状态</title>
</head>
<body background="../image/manager_bg.gif">

<table border="0" cellSpacing="0" width="279" background="../image/manager_bg.gif" height="416" oncontextmenu=self.event.returnValue=false>
<TBODY>
  <tr>
    <td colSpan="2" width="100%" height="40"><table border="0" cellSpacing="1" width="100%">
<TBODY>
      <tr>
        <td width="20" align="center"><a href="zt.asp" ><img alt="===========&#13&#10门派：<%=rs("门派")%> &#13&#10身份：<%=rs("身份")%> &#13&#10姓名：<%=rs("姓名")%>&#13&#10===========&#13&#10点击刷新本页&#13&#10==========="  src='<%=rs("头像")%>' border="0"></a></td>
        <td align="center"><%=rs("门派")%> [<%=rs("身份")%>] <b><font color="#dd2222"><%=rs("姓名")%></font></b> 个人资料<br></td>
      </tr>
</TBODY>
    </table>
    </td>
  </tr>
  <tr>
    <td colSpan="2" width="100%" height="79" style="border: 1px solid #FFFFFF"><div align="center"><center><table border="0"
    cellPadding="0" cellSpacing="0" width="100%">
<TBODY>
      <tr>
        <td align="center">攻击</td>
        <td align="center">防御</td>
        <td align="center">体力</td>
        <td align="center">内力</td>
        <td align="center">银两</td>
      </tr>
      <tr>
        <td align="center"><img height="46" src="../image/usergj.gif" width="47" alt="攻击:<%=rs("攻击")%>"></td>
        <td align="center"><img height="46" src="../image/userfy.gif" width="47" alt="防御:<%=rs("防御")%>"></td>
        <td align="center"><img height="46" src="../image/usersf.gif" width="47" alt="体力:<%=rs("体力")%>"></td>
        <td align="center"><img height="46" src="../image/usertl_.gif" width="47" alt="内力:<%=rs("内力")%>"></td>
        <td align="center"><img height="46" src="../image/useryz.gif" width="47" alt="银两:<%=rs("银两")%>"></td>
      </tr>
      <tr>
        <td align="center"><%=rs("攻击")%></td>
        <td align="center"><%=rs("防御")%></td>
        <td align="center"><%=rs("体力")%></td>
        <td align="center"><%=rs("内力")%></td>
        <td align="center"><%=rs("银两")%></td>
      </tr>
</TBODY>
    </table>
    </center></div></td>
  </tr>
  <tr>
    <td colSpan="2" height="15">
      <p align="center"><font color="#669900">鼠标放在图片上可以看到数值</font></td>
  </tr>
  <tr>
    <td height="19">&nbsp; 存款:<%=rs("存款")%> 两</td>                              
    <td height="19">伴侣:<%=rs("配偶")%></td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 积分:<%=rs("积分")%> 元</td>                              
    <td height="19">道德:<%=rs("道德")%></td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 资质:<%=rs("资质")%></td>                              
    <td height="19">情人:<%=rs("洞房花烛")%> </td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 精力:<%=rs("精力")%></td>                              
    <td height="19">等级:<%=rs("等级")%> 级</td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 荣誉称号:<%if rs("等级")="" then response.write("武林新手")                                                         
if rs("等级")<=3 then response.write("武林新手")                                                         
if rs("等级")<=6 and rs("等级")>3  then response.write("武林游侠")                                                         
if rs("等级")<=10 and rs("等级")>6  then response.write("武林斗侠")                                                         
if rs("等级")<=15and rs("等级")>10  then response.write("初级战将")                                                         
if rs("等级")<=20 and rs("等级")>15  then response.write("中级战将")                                                         
if rs("等级")<=25 and rs("等级")>20  then response.write("高级战将")                                                         
if rs("等级")<=30 and rs("等级")>25  then response.write("超级战将")                        
if rs("等级")<=35 and rs("等级")>30  then response.write("武神")                                                           
if rs("等级")<38 and rs("等级")>35 then response.write("战神")                         
if rs("等级")>=38 then response.write("终极战神")                                                         
%></td>                              
    <td height="19">小妾:<%=rs("小妾")%></td>                              
  </tr>                              
</TBODY>                              
  <tr>                              
    <td height="19">&nbsp; 江湖名号:<%=rs("头衔")%></td>                              
    <td height="19">介绍人:<%=rs("介绍人")%></td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 门派:<%=rs("门派")%></td>                            
    <td height="19">身份:<%=rs("身份")%></td>                            
  </tr>                            
  <tr>                            
    <td height="19">&nbsp; 国家:<%=rs("国家")%></td>                           
    <td height="19">职位:<%=rs("职位")%></td>                           
  </tr>                           
  <tr>                             
    <td height="19">&nbsp; 职业:<%=rs("职业")%> </td>                              
    <td height="19">师傅:<%=rs("师傅")%></td>                              
  </tr>                              
  <tr>                              
    <td height="19">&nbsp; 被杀:<%=rs("被杀")%>次</td>                             
    <td height="19">杀人:<%=rs("总杀人")%>个</td>                             
  </tr>                             
  <tr>                             
    <td height="19">&nbsp; 结拜:<%=rs("结拜兄弟")%></td>                             
    <td height="19">会员期:<%=rs("会员时间")%></td>                            
  </tr>                            
  <tr>                            
    <td width="100%" colspan="2" height="46"><table border="0" width="100%" cellspacing="0" cellpadding="0">                            
  <tr>                            
    <td width="100%" align="center" height="30">个人签名</td>                            
  </tr>                            
  <tr>                            
    <td width="100%"><%=rs("签名档")%></td>                            
 </tr>                            
</table></td>                            
  </tr>                            
</table>                            
<%                            
rs.close                            
set rs=nothing                            
conn.close                            
set conn=nothing                            
%>