<%@ LANGUAGE=VBScript codepage ="936" %>
<%aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="Select  * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.Execute(sql)
%>
<link href="../hc3w_img/hc3w_main.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF">
<table border="0" cellSpacing="0" width="279" background="../hc3w_img/manager_bg.gif" height="388">
<TBODY>
  <tr>
    <td colSpan="2" width="100%" height="42"><table border="0" cellSpacing="1" width="100%">
<TBODY>
      <tr>
        <td width="20" align="center"><a href="sm.asp" target="gr"><img alt="===========&#13&#10门派：<%=rs("门派")%> &#13&#10身份：<%=rs("身份")%> &#13&#10姓名：<%=rs("姓名")%>&#13&#10===========&#13&#10   点击刷新本页&#13&#10==========="  src='../ico/<%=rs("名单头像")%>-2.gif' border="0"></a></td>
        <td align="center"><%=rs("门派")%> [<%=rs("身份")%>] <b><font color="#dd2222"><%=rs("姓名")%></font></b> 个人资料<br></td>
      </tr>
</TBODY>
    </table>
    </td>
  </tr>
  <tr>
    <td colSpan="2" width="100%" height="81"><div align="center"><center><table border="0"
    cellPadding="0" cellSpacing="0" width="100%">
<TBODY>
      <tr>
        <td align="center">攻击</td>
        <td align="center">防御</td>
        <td align="center">武功</td>
        <td align="center">内功</td>
        <td align="center">银两</td>
      </tr>
      <tr>
        <td align="center"><img height="46" src="../hc3w_img/usergj.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/userfy.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/usersf.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/usertl_.gif" width="47"></td>
        <td align="center"><img height="46" src="../hc3w_img/useryz.gif" width="47"></td>
      </tr>
      <tr>
        <td align="center"><%=rs("攻击")%></td>
        <td align="center"><%=rs("防御")%></td>
        <td align="center"><%=rs("武功")%></td>
        <td align="center"><%=rs("内力")%></td>
        <td align="center"><%=rs("银两")%></td>
      </tr>
</TBODY>
    </table>
    </center></div></td>
  </tr>
  <tr>
    <td colSpan="2" height="17"></td>
  </tr>
  <tr>
    <td height="17">&nbsp; 存&nbsp; 款：<%=rs("存款")%> 两</td>
    <td height="17">伴&nbsp; 侣：<%=rs("配偶")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; 会&nbsp; 费：<%=rs("会员金卡")%> 元</td>
    <td height="17">道&nbsp; 德：<%=rs("道德")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; 魅&nbsp; 力：<%=rs("魅力")%></td>
    <td height="17">赢/赌：<%=rs("赢次数")%> / <%=rs("赌次数")%></td>
  </tr>
  <tr>
    <td height="17">&nbsp; 体&nbsp; 力：<%=rs("体力")%></td>
    <td height="17">战斗级：<%=rs("等级")%> 级</td>
  </tr>
  <tr>
    <td height="6">&nbsp; 赌场等级：<%if rs("赢钱")<100  then response.write("赌场菜鸟")
if rs("赢钱")>=2000 and rs("赢钱")<10000  then response.write("赌场游民")
if rs("赢钱")>=10000 and rs("赢钱")<50000  then response.write("初级赌徒")
if rs("赢钱")>=50000 and rs("赢钱")<100000  then response.write("赌场游侠")
if rs("赢钱")>=100000 and rs("赢钱")<300000  then response.write("职业高手")
if rs("赢钱")>=300000 and rs("赢钱")<800000  then response.write("赌侠")
if rs("赢钱")>=800000 and rs("赢钱")<1500000  then response.write("赌王")
if rs("赢钱")>=1000000 then response.write("赌神")
%></td>
    <td height="6">管理级：<%=rs("grade")%> 级</td>
  </tr>
</TBODY>
  <tr>
    <td height="6">&nbsp; 江湖名号：<%if rs("等级")<5  then response.write("初来乍道")
if rs("等级")>=5 and rs("等级")<10  then response.write("江湖初行")
if rs("等级")>=10 and rs("等级")<15  then response.write("小有成就")
if rs("等级")>=15 and rs("等级")<20  then response.write("声名显赫")
if rs("等级")>=20 and rs("等级")<35  then response.write("行闯江湖")
if rs("等级")>=35 and rs("等级")<45  then response.write("一代大侠")
if rs("等级")>=45 and rs("等级")<55  then response.write("江湖剑客")
if rs("等级")>=55 and rs("等级")<65  then response.write("闻名天下")
if rs("等级")>=65 and rs("等级")<75  then response.write("云游仙胜")
if rs("等级")>=100 then response.write("剑仙")
%></td>
    <td height="6">介绍人：<%=rs("介绍人")%></td>
  </tr>
  <tr>
    <td height="5">&nbsp; E-MAIL：<%=rs("信箱")%></td>
    <td height="5">OICQ：<%=rs("oicq")%></td>
  </tr>
  <tr>
    <td height="5">&nbsp; 职&nbsp; 业：<%=rs("职业")%> 宝物:<%=rs("宝物")%></td>
    <td height="5">师&nbsp; 傅：<%=rs("师傅")%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2"><table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" align="center" height="30">个人签名</td>
  </tr>
  <tr>
    <td width="100%"><%=rs("签名")%></td>
  </tr>
</table></td>
  </tr>
</table>

<p>
  <%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</p>