<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
action=trim(request.querystring("action"))
if action="大家" then action=aqjh_name
if aqjh_name<>action then
	rs.open "Select 银两,等级,grade,种族,威望,国家,职位 from 用户 where 姓名='"&aqjh_name&"'",conn
	if rs("银两")<10000 and rs("grade")<7 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你没有一万两银子打点查不到东西的！');window.close();</script>"
		response.end
	end if
	dj=rs("等级")
	conn.Execute "update 用户 set 银两=银两-10000 where 姓名='"&aqjh_name&"'"
rs.close
end if
rs.open "Select  * from 用户 where 姓名='"& action &"'",conn
%>
<html>
<head>
<title>江湖档案</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script language=javascript>window.moveTo(100,50);</script>
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: 宋体 }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body style="BACKGROUND-COLOR: buttonface; BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; BORDER-RIGHT: medium none; BORDER-TOP: medium none; PADDING-BOTTOM: 6px; PADDING-LEFT: 6px; PADDING-RIGHT: 6px; PADDING-TOP: 6px" marginwidth="0" marginheight="0" leftmargin="0" topmargin="0" LANGUAGE="javascript"  >
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center">
  <tr bgcolor="#000000">
<td colspan="8" height="33">
<div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF">ID:</font><font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("id")%>
</font></b></font></font><font color="#FFFFFF"><%=rs("姓名")%></font></b><font
size="2" color="#FFFFFF">大侠的江湖背景</font></font></font> </div>
</td>
</tr>
<tr >
<td width="36" height="12"><font size="2" class="c">姓 名:</font></td>
<td colspan="2" height="12"><%=rs("姓名")%></td>
<td height="12" colspan="2">地区:<%=rs("地区")%></td>
<td height="12" width="49">
<div align="center">名 号:</div>
</td>
<td height="12" colspan="2"> <font color="#0000FF" size="2">
<%
if aqjh_name=action then
zddj=-1
else
zddj=rs("等级")
end if
if rs("等级")<5  then response.write("初来乍道")
if rs("等级")>=5 and rs("等级")<10  then response.write("江湖初行")
if rs("等级")>=10 and rs("等级")<15  then response.write("小有成就")
if rs("等级")>=15 and rs("等级")<20  then response.write("声名显赫")
if rs("等级")>=20 and rs("等级")<35  then response.write("行闯江湖")
if rs("等级")>=35 and rs("等级")<45  then response.write("一代大侠")
if rs("等级")>=45 and rs("等级")<55  then response.write("江湖剑客")
if rs("等级")>=55 and rs("等级")<65  then response.write("闻名天下")
if rs("等级")>=65 and rs("等级")<75  then response.write("云游仙胜")
if rs("等级")>=75 and rs("等级")<80  then response.write("已入仙道")
if rs("等级")>=80 and rs("等级")<85  then response.write("小仙")
if rs("等级")>=85 and rs("等级")<90  then response.write("大仙")
if rs("等级")>=90 and rs("等级")<95  then response.write("极帝大仙")
if rs("等级")>=95 and rs("等级")<100  then response.write("仙人")
if rs("等级")>=100 then response.write("剑仙")
%>
</font></td>
</tr>
<tr >
<td width="36" height="14"><font size="2" class="c">性 别:</font></td>
<td height="14" width="33"><font color="#000000" size="2">
<%if rs("性别")="男"  then response.write("帅哥")
if rs("性别")="女"  then response.write("美女")
%>
</font></td>
<td height="14" width="37">年龄:</td>
<td height="14" nowrap width="96"><%=rs("年龄")%> 岁</td>
<td height="14" width="41">状态:</td>
<td height="14" width="49"><%=rs("状态")%></td>
<td height="14" width="34">师傅:</td>
<td height="14" width="69"><%=rs("师傅")%></td>
</tr>
<tr >
<td width="36"><font size="2" class="c">Email:</font></td>
<td colspan="3" nowrap><%=rs("信箱")%> </td>
<td width="41">OIcq:</td>
<td width="49"><%=rs("oicq")%></td>
<td width="34">职业:</td>
<td width="69"><%=rs("职业")%></td>
</tr>
</table>
<div align="left"></div>
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center">
  <tr > 
    <td colspan="5" height="20"> 
      <div align="center"> <font size="2">江 湖 档 案</font> </div>
    </td>
  </tr>
  <tr > 
    <td width="74" height="25">现 金：</td>
    <td height="25" width="134"> 
      <%if dj>zddj then%>
      <%=rs("银两")%> 两 
      <%else%>
      等级不够 
      <%end if%>
    </td>
    <td width="74" height="25">介绍人：</td>
    <td height="25" width="47"> <%=rs("介绍人")%> </td>
    <td height="25" width="93">金卡： 
      <%if dj>zddj then%>
      <font color="#0000FF"><%=rs("会员金卡")%>元</font> 
      <%else%>
      等级不够 
      <%end if%>
    </td>
  </tr>
  <tr > 
    <td width="74" height="20">存 款：</td>
    <td height="24" width="134"> 
      <%if dj>zddj then%>
      <%=rs("存款")%> 两 
      <%else%>
      等级不够 
      <%end if%>
    </td>
    <td width="74" height="24">道 德：</td>
    <td height="24" colspan="2"> <%=rs("道德")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">武 功：</td>
    <td height="24" width="134"> <%=rs("武功")%></td>
    <td width="74" height="24">攻 击：</td>
    <td height="24" colspan="2"> <%=rs("攻击")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">内 力：</td>
    <td width="134"> <%=rs("内力")%> </td>
    <td width="74">防 御：</td>
    <td colspan="2"> <%=rs("防御")%> </td>
  </tr>
  <tr > 
    <td width="74" height="20">魅 力：</td>
    <td width="134"> <%=rs("魅力")%> </td>
    <td width="74">门 派：</td>
    <td width="47"><%=rs("门派")%> </td>
    <td width="93">宝物:<%=rs("宝物")%></td>
  </tr>
  <tr > 
    <td width="74" height="20">体 力：</td>
    <td width="134"> <%=rs("体力")%> </td>
    <td width="74">身 份：</td>
    <td colspan="2"> <%=rs("身份")%> </td>
  </tr>
<tr > 
    <td width="74" height="20">威 望：</td>
    <td width="134"> <%=rs("威望")%> </td>
    <td width="74">种 族：</td>
    <td colspan="2"> <%=rs("种族")%> </td>
  </tr>
<tr > 
    <td width="74" height="20">国 家：</td>
    <td width="134"> <%=rs("国家")%> </td>
    <td width="74">职 位：</td>
    <td colspan="2"> <%=rs("职位")%> </td>
  </tr>
  <tr > 
    <td width="74" height="27">配 偶：</td>
    <td width="134" height="27"> <%=rs("配偶")%> <font color="#0000FF">结婚次数:<%=rs("结婚次数")%></font></td>
    <td width="74" height="27">赢 / 赌：</td>
    <td colspan="2" height="27"><%=rs("赢次数")%> / <%=rs("赌次数")%></td>
  </tr>
  <tr > 
    <td width="74" height="20">赌场等级：</td>
    <td width="134"> 
      <%if rs("赢钱")<100  then response.write("赌场菜鸟")
if rs("赢钱")>=2000 and rs("赢钱")<10000  then response.write("赌场游民")
if rs("赢钱")>=10000 and rs("赢钱")<50000  then response.write("初级赌徒")
if rs("赢钱")>=50000 and rs("赢钱")<100000  then response.write("赌场游侠")
if rs("赢钱")>=100000 and rs("赢钱")<300000  then response.write("职业高手")
if rs("赢钱")>=300000 and rs("赢钱")<800000  then response.write("赌侠")
if rs("赢钱")>=800000 and rs("赢钱")<1500000  then response.write("赌王")
if rs("赢钱")>=1000000 then response.write("赌神")
%>
    </td>
    <td width="74">江湖等级：</td>
    <td width="47"><%=rs("等级")%>级</td>
    <td width="93">管理等级:<%=rs("grade")%>级</td>
  </tr>
  <tr > 
    <td colspan="5" height="199"> 
      <table width="424" border="1" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="196" width="223"> 
         <div align="center">
<IFRAME src='../show/userface.asp?username=<%=rs("姓名")%>&sex=<%=rs("性别")%>' width="148" height="233" leftmargin='0' marginWidth='0' marginHeight='0' scrolling='no'></IFRAME></div>
          </td>
          <td height="196" width="184" align="left" valign="top">个人签名：<span class="txt"><%=changechr(rs("签名"))%></span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body></html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%>