<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("sjjh_usermdb")
action=sjjh_name
sql="Select  * from 用户 where 姓名='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>

<head>
<title>江湖档案修改♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: 宋体 }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body bgcolor="#BABABA" topmargin="0" leftmargin="0">
<form method="POST" action="regmodi1.asp">
<table border="1" width="429" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
<tr bgcolor="#000000">
<td colspan="8" height="33">
<div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("姓名")%></font></b><font
size="2" color="#FFFFFF">大侠的江湖背景</font></font></font> </div>
</td>
</tr>
<tr bgcolor="#cccccc">
<td width="36" height="12"><font size="2" class="c">姓 名:</font></td>
<td colspan="2" height="12"><%=rs("姓名")%></td>
<td height="12" colspan="2">地区:
<input type="text"
name="diqu" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("地区")%>">
</td>
<td height="12" width="66">
<div align="center">名 号:</div>
</td>
<td height="12" colspan="2"> <font color="#0000FF" size="2">
<%if rs("等级")<5  then response.write("初来乍道")
if rs("等级")>=5 and rs("等级")<10  then response.write("江湖初行")
if rs("等级")>=10 and rs("等级")<15  then response.write("小有成就")
if rs("等级")>=15 and rs("等级")<20  then response.write("声名显赫")
if rs("等级")>=20 and rs("等级")<35  then response.write("行闯江湖")
if rs("等级")>=35 and rs("等级")<45  then response.write("一代大侠")
if rs("等级")>=45 and rs("等级")<55  then response.write("江湖剑客")
if rs("等级")>=55 and rs("等级")<65  then response.write("闻名天下")
if rs("等级")>=65 and rs("等级")<75  then response.write("云游仙胜")
if rs("等级")>=100 then response.write("剑仙")
%>
</font></td>
</tr>
<tr bgcolor="#cccccc">
<td width="36" height="20"><font size="2" class="c">性 别:</font></td>
<td height="20" width="38"><font color="#000000" size="2">
<%if rs("性别")="男"  then response.write("帅哥")
if rs("性别")="女"  then response.write("美女")
%>
</font></td>
<td height="20" width="35">年龄:</td>
<td height="20" nowrap width="71">
<input type="text"
name="nianling" size="2"
style="font-family: Tahoma; font-size: 12px"
maxlength="3" value="<%=rs("年龄")%>">
岁 </td>
<td height="20" width="40">状态:</td>
<td height="20" width="66"><%=rs("状态")%></td>
<td height="20" width="36">师傅:</td>
<td height="20" width="73"><%=rs("师傅")%></td>
</tr>
<tr bgcolor="#cccccc">
<td width="36"><font size="2" class="c">Email:</font></td>
<td colspan="3" nowrap>
<input type="text"
name="email" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("信箱")%>">
</td>
<td width="40">OIcq:</td>
<td width="66">
<input type="text"
name="oicq" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9" value="<%=rs("oicq")%>">
</td>
<td width="36">职业:</td>
<td width="73"><%=rs("职业")%></td>
</tr>
</table>
<div align="left"></div>
  <table border="1"
width="424" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
    <tr bgcolor="#cccccc"> 
      <td colspan="5" height="20"> 
        <div align="center"> <font size="2">江 湖 档 案</font> </div>
      </td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="2">现 金：</td>
      <td width="169" height="2"><%=rs("银两")%> 两</td>
      <td width="56" height="2">介绍人：</td>
      <td height="2" width="50"><%=rs("介绍人")%></td>
      <td height="2" width="84">会员费：<%=rs("会员金卡")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">存 款：</td>
      <td width="169" height="24"><%=rs("存款")%> 两</td>
      <td width="56" height="24">道 德 ：</td>
      <td height="24" colspan="2"><%=rs("道德")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">武 功：</td>
      <td width="169" height="24"><%=rs("武功")%></td>
      <td width="56" height="24">攻 击 ：</td>
      <td height="24" colspan="2"><%=rs("攻击")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">内 力：</td>
      <td width="169"><%=rs("内力")%></td>
      <td width="56">防 御 ：</td>
      <td colspan="2"><%=rs("防御")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">魅 力：</td>
      <td width="169"><%=rs("魅力")%></td>
      <td width="56">门 派 ：</td>
      <td colspan="2"><%=rs("门派")%> </td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">体 力：</td>
      <td width="169"><%=rs("体力")%></td>
      <td width="56">身 份 ：</td>
      <td colspan="2"><%=rs("身份")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="20">配 偶：</td>
      <td width="169"><%=rs("配偶")%></td>
      <td width="56">赢 / 赌：</td>
      <td colspan="2"><%=rs("赢次数")%> / <%=rs("赌次数")%></td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td width="52" height="2">赌 场：</td>
      <td width="169" height="2"> 
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
      <td width="56" height="2">等 级：</td>
      <td width="50" height="2"><%=rs("等级")%>级</td>
      <td width="84" height="2">管理级:<%=rs("grade")%>级</td>
    </tr>
    <tr bgcolor="#cccccc"> 
      <td colspan="5" height="199"> 
        <table width="421" border="1" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="196" width="237"> 
              <div align="center"><font color="#000000" size="2"> </font> <font color="#000000"><img src="<%=rs("头像")%>"></font><br>
                <br>
                头像url地址： <br>
                <input type="text"
name="touxiang" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="100" value="<%=rs("头像")%>">
                <br>
                输入头像的url地址，可以修改自己<br>
                的头像，您可以用网易、逸飞岭中的<br>
                头像来作自己的，头像不要太大：建议<br>
                采用：200*130大小 </div>
            </td>
            <td height="196" width="178" align="left" valign="top"> 
              <div align="center">个人签名，写上想与朋友说的话！<span class="txt"> 
                <textarea name="qianming" cols="30" style="font-family: Tahoma; font-size: 12px" rows="20"><%=rs("签名")%></textarea>
                </span></div>
            </td>
          </tr>
        </table>
        <div align="right"></div>
        <div align="right"></div>
      </td>
    </tr>
  </table>
<div align="center"><br>
<input type="submit" value="江湖资料修改" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</div>
</form>
<div align="center">
<br>
</div>
</body>

</html>
<%
rs.close
set rs=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%>