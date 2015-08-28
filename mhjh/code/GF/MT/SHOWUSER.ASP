<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
myName=session("yx8_mhjh_username")
if myName="" then Response.Redirect "../../error.asp?id=016"
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next

if check<>Application("yx8_mhjh_banbenhao") then session.Abandon
username=Request.Form("search")
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT 银两 FROM 用户 where 姓名='"&myName&"' and 银两>1000"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('你身上的银两不够阿！回去努力吧');history.back();</script>"
conn.close
Response.end
else
sqlstr="SELECT * FROM 用户 where 姓名='"&username&"'"
Set Rs=conn.Execute(sqlstr)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('查无此人');history.back();</script>"
conn.close
Response.end
else
%>
<html>

<head>
<title>这是你要查的资料，多谢你的1000两了。</title>

<link rel=stylesheet href='../../style.css'>
</head>

<body oncontextmenu=self.event.returnValue=false background="../../image/manager_bg.gif" text="#FFFFFF">
<table border="1" cellspacing="1" cellpadding="0" align="center"
width="429" bordercolor="#000000">
<tr >
<td colspan="2">
<div align="center"> <font size="3" ><b><font color="#FFFFFF"><%=rs("姓名")%></font></b><font
size="2" color="#FFFFFF">大侠的</font></font><font size="2" color="#FFFFFF" class="c">江湖背景</font> </div>
</td>
</tr>
<tr>
<td width="64"><font class="c">姓 名</font></td>
<td width="219"><%=rs("姓名")%></td>
</tr>
<tr>
<td width="64"><font class="c">性 别</font></td>
<td width="219"><%=rs("性别")%> </td>
</tr>
<tr>
<td width="64"><font class="c">Email地址</font></td>
<td width="219"><%=rs("电子邮箱")%> </td>
</tr>
</table>
<table border="1" cellspacing="1" cellpadding="0" align="center"
width="429" bordercolor="#000000">
<tr>
<td colspan="4" height="20">
<div align="center"> <font size="2">剑 侠 档 案</font> </div>
</td>
</tr>
<tr>
<td width="64" height="25">现 金：</td>
<td width="141" height="25"><%=rs("银两")%> 两</td>
<td width="66" height="25">精 力：</td>
<td width="127" height="25"><%=rs("精力")%></td>
</tr>
<tr>
<td width="64" height="20">存 款：</td>
<td width="141" height="24">保密</td>
<td width="66" height="24">道 德：</td>
<td width="127" height="24"><%=rs("道德")%></td>
</tr>
<tr>
<td width="64" height="20">积 分：</td>
<td width="141" height="24"><%=rs("积分")%></td>
<td width="66" height="24">攻 击：</td>
<td width="127" height="24"><%=rs("攻击")%></td>
</tr>
<tr>
<td width="64" height="20">内 力：</td>
<td width="141"><%=rs("内力")%></td>
<td width="66">防 御：</td>
<td width="127"><%=rs("防御")%></td>
</tr>
<tr>
<td width="64" height="20">资 质：</td>
<td width="141"><%=rs("资质")%></td>
<td width="66">门 派：</td>
<td width="127"><%=rs("门派")%> </td>
</tr>
<tr>
<td width="64" height="20">体 力：</td>
<td width="141"><%=rs("体力")%></td>
<td width="66">身 份：</td>
<td width="127"><%=rs("身份")%></td>
</tr>
<tr>
<td width="64" height="20">配 偶：</td>
<td width="141"><%=rs("配偶")%></td>
<td width="66">等 级：</td>
<td width="127"><%=rs("等级")%> </td>
</tr>
<tr>
<td width="64" height="20">国 家：</td>
<td width="141"><%=rs("国家")%></td>
<td width="66">职 位：</td>
<td width="127"><%=rs("职位")%></td>
</tr>
<tr>
<td width="64" height="20">状 态：</td>
<td width="141"><%=rs("状态")%>
</td>
<td width="66">荣 誉：</td>
<td width="127"><%if rs("等级")="" then response.write("武林新手")                                                          
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
</tr>
</table>
<div align="center"><br>
<%
sql="update 用户 set 银两=银两-1000 where 姓名='"&myName&"'"
conn.Execute(sql)
%>
<a href="index.asp">返回密探住所</a>&nbsp;&nbsp;</div>
</html>
<%
end if
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
