<%@LANGUAGE="VBSCRIPT" CodePage="936"%>
<%
'happyjh.com
Option Explicit
Response.Buffer = True
Server.scriptTimeout="20"
public const mpassword="1997" '设置管理的密码
dim starttime:starttime=timer()

function leftshow(str,leftc)
if len(str)>=leftc then
leftshow=left(str,leftc)&".."
else
leftshow=str
end if
end function

function htmlencode(reString)
dim Str
str=reString
str=replace(str, "&", "&amp;")
str=replace(str, ">", "&gt;")
str=replace(str, "<", "&lt;")
htmlencode=Str
end function

sub cklogin()
if not request.cookies("jqqonline")("password")=mpassword then
response.clear
response.write "Error !!"
response.end
end if
end sub

sub login()
if request.cookies("jqqonline")("error")="2" then
%>
<script language="javascript">
alert("您已经输入了两次错误的密码\n\n请关闭浏览器以后重新尝试")
history.back()
</script>
<%
response.end
end if
if request.form("password")=mpassword then
response.cookies("jqqonline")("password")=request.form("password")
response.redirect "?type=manage"
else
if request.cookies("jqqonline")("error")="" then
response.cookies("jqqonline")("error")="1"
else
response.cookies("jqqonline")("error")="2"
end if
%>
<script language="javascript">
alert("密码错误！")
history.back()
</script>
<%
end if
response.end
end sub


sub editsiteinfo()
Dim infostrSourceFile,infoobjXML
infostrSourceFile=Server.MapPath("xml/info.xml")
Set infoobjXML=Server.CreateObject("Microsoft.XMLDOM")
infoobjXML.load(infostrSourceFile)
Dim infoobjNodes
Set infoobjNodes=infoobjXML.selectSingleNode("xml/qqinfo/qqset[siteid ='1']")
If Not IsNull(infoobjNodes) then
infoobjNodes.childNodes(0).text=htmlencode(request.form("sitename"))
infoobjNodes.childNodes(1).text=htmlencode(request.form("siteskin"))
infoobjNodes.childNodes(2).text=htmlencode(request.form("siteshowx"))
infoobjNodes.childNodes(3).text=htmlencode(request.form("siteshowy"))
infoobjXML.save(infostrSourceFile)
%>
<script language="javascript">
alert("参数修改成功！")
location.href="?type=manage"
</script>
<%
Else
%>
<script language="javascript">
alert("Xml 未成功打开！")
history.back()
</script>
<%
End If
Set infoobjNodes=nothing
Set infoobjXML=nothing
end sub


sub addinfo()
if trim(request.form("qq"))="" or trim(request.form("dis"))="" or trim(request.form("face"))="" then
%>
<script language="javascript">
alert("没有填入必要的数据！")
history.back()
</script>
<%
response.end
end if

dim jtb_color
if trim(request.form("color"))="" then
jtb_color="#000000"
else
jtb_color=htmlencode(request.form("color"))
end if

Dim strSourceFile,objXML,oListNode,oDetailsNode,AllNodesNum
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Dim objRootlist
Set objRootlist=objXML.documentElement.selectSingleNode("qqlist")
dim id
If objRootlist.hasChildNodes then
id=objRootlist.lastChild.lastChild.text+1
Else
id=1
End If
Set objRootlist=nothing
Set oListNode=objXML.documentElement.selectSingleNode("qqlist").AppendChild(objXML.createElement("qq"))
Set oDetailsNode=oListNode.appendChild(objXML.createElement("qid"))
oDetailsNode.Text=htmlencode(request.form("qq"))
Set oDetailsNode=oListNode.appendChild(objXML.createElement("dis"))
oDetailsNode.Text=htmlencode(request.form("dis"))
Set oDetailsNode=oListNode.appendChild(objXML.createElement("face"))
oDetailsNode.Text=htmlencode(request.form("face"))
Set oDetailsNode=oListNode.appendChild(objXML.createElement("color"))
oDetailsNode.Text=jtb_color
Set oDetailsNode=oListNode.appendChild(objXML.createElement("id"))
oDetailsNode.Text=id
objXML.save(strSourceFile)
Set objRootlist=nothing
Set oListNode=nothing
Set oDetailsNode=nothing
Set objXML=nothing
%>
<script language="javascript">
alert("添加新的QQ号成功！")
location.href="?type=manage"
</script>
<%
response.end
end sub



sub editinfo()
dim editid:editid=request.querystring("id")
if not IsNumeric(editid) or editid="" then
%>
<script language="javascript">
alert("非法操作！")
history.back()
</script>
<%
response.end
else
editid=clng(editid)
end if
if trim(request.form("qq"))="" or trim(request.form("dis"))="" or trim(request.form("face"))="" then
%>
<script language="javascript">
alert("没有添入必要的数据！")
history.back()
</script>
<%
response.end
end if
Dim strSourceFile,objXML
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile) 
Dim objNodes
Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&editid&"']")
If Not IsNull(objNodes) then
objNodes.childNodes(0).text=htmlencode(request.form("qq"))
objNodes.childNodes(1).text=htmlencode(request.form("dis"))
objNodes.childNodes(2).text=htmlencode(request.form("face"))
objNodes.childNodes(3).text=htmlencode(request.form("color"))
objXML.save(strSourceFile)
Set objNodes=nothing
Set objXML=nothing
%>
<script language="javascript">
alert("修改成功！")
location.href="?type=manage"
</script>
<%
else
%>
<script language="javascript">
alert("修改失败！")
location.href="?type=manage"
</script>
<%
end if
response.end
end sub

sub delinfo()
dim delid
delid=request.querystring("id")
if not IsNumeric(delid) or delid="" then
%>
<script language="javascript">
alert("非法操作！")
location.href="<%=jurl%>?type=manage"
</script>
<%
response.end
else
delid=clng(delid)
end if
Dim strSourceFile,objXML
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Dim objNodes
Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&delid&"']")
if Not IsNull(objNodes) then
if request.querystring("yn")="" then
%>
<script language="javascript">
if(confirm("确认要删除[<%=objNodes.childNodes(0).text%>]的信息吗？"))
window.location="?type=manage&act=delete&id=<%=delid%>&yn=1"
else
history.back()
</script>
<%
else
objNodes.parentNode.removeChild(objNodes)
objXML.save(strSourceFile)
%>
<script language="javascript">
alert("删除成功！")
location.href="?type=manage"
</script>
<%
end if
else
%>
<script language="javascript">
alert("没有找到指定的条目！")
location.href="?type=manage"
</script>
<%
End If
Set objNodes=nothing
Set objXML=nothing
response.end
end sub


select case request.querystring("act")
case "login" call login()
case "editsiteinfo" call cklogin():call editsiteinfo()
case "add" call cklogin():call addinfo()
case "edit" call cklogin():call editinfo()
case "delete" call cklogin():call delinfo()
end select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>JQQonline插件信息管理界面</title>
<style type="text/css">
<!--
body{margin-left: 0px;margin-top: 0px;margin-right: 0px;margin-bottom: 0px;BACKGROUND:#C7E3FE;}
a{font-style:normal;TEXT-DECORATION: none;color:#000000;}
a:hover{font-style:normal;TEXT-DECORATION: none;color:#F4F4F4;}
a:active{font-style:normal;TEXT-DECORATION: none;color:#000000;}
table{font-size: 9pt;font-family: tahoma;color:#000000;}
-->
</style>
</head>
<body>
<%
sub nlogin()
%>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500" height="20"></td>
  </tr>
</table>
<form method="post" action="?act=login">
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500">
    <fieldset align="center">
    <legend>管理员登陆</legend>    
    <table border="0" width="500">
      <tr>
        <td width="500" height="10"></td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;管理密码：<input type="password" name="password" size="30">&nbsp;<input type="submit" name="submit" value="登陆"></td>
      </tr>     
      <tr>
        <td width="500" height="20"></td>
      </tr>
    </table>
    </fieldset>
    </td>
  </tr>
</table>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500" height="10"></td>
  </tr>
</table>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500" height="20">通用网站在线咨询JQQonline插件 V3.0 <a href="http://www.happyjh.com" target="_blank">happyjh.Com</a></td>                                                                           
  </tr>
</table>
</form>
<%
end sub

if request.querystring("action")="login" then '设置登陆方式，该模式为在文件名后输入?action=login进入登陆页面
call nlogin()
response.write "</body>"
response.write "</html>"
response.end
end if


sub slist()
%>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500">
    <fieldset align="center">
    <legend>通用网站在线咨询JQQonline插件</legend>
<%
Dim infostrSourceFile,infoobjXML
infostrSourceFile=Server.MapPath("xml/info.xml")
Set infoobjXML=Server.CreateObject("Microsoft.XMLDOM")
infoobjXML.load(infostrSourceFile)
Dim infoobjNodes
Set infoobjNodes=infoobjXML.selectSingleNode("xml/qqinfo/qqset[siteid ='1']")
If Not IsNull(infoobjNodes) then
%>
    <form method="post" action="?type=manage&act=editsiteinfo">
    <table border="0" width="500">    
      <tr>
        <td width="500" height="25">&nbsp;参数设置</td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;网站名：<input type="text" name="sitename" size="15" value="<%=infoobjNodes.childNodes(0).text%>">&nbsp;使用皮肤：<input type="text" name="siteskin" size="5" value="<%=infoobjNodes.childNodes(1).text%>"></td>
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;显示界面X坐标：<input type="text" name="siteshowx" size="10" value="<%=infoobjNodes.childNodes(2).text%>">&nbsp;Y坐标：<input type="text" name="siteshowy" size="10" value="<%=infoobjNodes.childNodes(3).text%>">&nbsp;<input type="submit" name="submit" value="修改"></td>
      </tr> 
      <tr>
        <td width="500" height="10"></td>
      </tr>
    </table>
    </form>
<%
Else
%>
<script language="javascript">
alert("Xml 未成功打开！")
history.back()
</script>
<%
response.end
End If
Set infoobjNodes=nothing
Set infoobjXML=nothing
%>
    <table border="0" width="500">
      <tr>
        <td width="180" height="25">&nbsp;QQ号</td>
        <td width="180" height="25">&nbsp;描述</td>
        <td width="80" height="25">&nbsp;头像</td>
        <td width="30" align="center" height="25">编辑</td>
        <td width="30" align="center" height="25">删除</td>
      </tr>
<%
Dim strSourceFile,objXML,objRootsite,AllNodesNum
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Set objRootsite=objXML.documentElement.selectSingleNode("qqlist")
AllNodesNum=objRootsite.childNodes.length-1
Dim iCount
For iCount=0 to AllNodesNum
%>
      <tr>
        <td width="180" height="25">&nbsp;<%=objRootsite.childNodes.item(iCount).childNodes.item(0).text%></td>
        <td width="180" height="25">&nbsp;<%=objRootsite.childNodes.item(iCount).childNodes.item(1).text%></td>
        <td width="80" height="25">&nbsp;<img src="images/qqface/<%=objRootsite.childNodes.item(iCount).childNodes.item(2).text%>_m.gif" border="0"></td>
        <td width="30" align="center" height="25"><a href="?type=manage&mtype=edit&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">编辑</a></td>
        <td width="30" align="center" height="25"><a href="?type=manage&act=delete&id=<%=objRootsite.childNodes.item(iCount).childNodes.item(4).text%>">删除</a></td>
      </tr>
<%
Next
Set objRootsite=nothing
Set objXML=nothing
%>      
    </table>
    <form method="post" action="?type=manage&act=add">
    <table border="0" width="500">
      <tr>
        <td width="500" height="10"></td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;添加新的QQ号</td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;QQ号：<input type="text" name="qq" size="20"></td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;描述：<input type="text" name="dis" size="25"></td>
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;颜色：<input type="text" name="color" size="25"> 输入颜色代码例如：#000000</td>      
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;头像：<br>   
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>"><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                                          
        <%next%>                                                          
        </td>                                                                           
      </tr>
      <tr>
        <td width="500" height="10"></td>
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;<input type="submit" name="submit" value="确定添加">&nbsp;<input type="reset" name="reset" value="取消重置"></td>
      </tr>
      <tr>
        <td width="500" height="10"></td>
      </tr>
    </table>
    </form>
    </fieldset>
    </td>
  </tr>
</table>
<%
end sub



sub neditinfo()
%>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500">
    <fieldset align="center">
    <legend>编辑QQ信息</legend>    
<%
dim neditid:neditid=request.querystring("id")
if not IsNumeric(neditid) or neditid="" then
%>
<script language="javascript">
alert("非法操作！")
history.back()
</script>
<%
response.end
else
neditid=clng(neditid)
end if
Dim strSourceFile,objXML
strSourceFile=Server.MapPath("xml/qq.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Dim objNodes
Set objNodes=objXML.selectSingleNode("xml/qqlist/qq[id ='"&neditid&"']")
If Not IsNull(objNodes) then
%>
    <form method="post" action="?type=manage&act=edit&id=<%=neditid%>">
    <table border="0" width="500">
      <tr>
        <td width="500" height="10"></td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;修改指定的QQ信息</td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;QQ号：<input type="text" name="qq" size="20" value="<%=objNodes.childNodes(0).text%>"></td>
      </tr>     
      <tr>
        <td width="500" height="25">&nbsp;描述：<input type="text" name="dis" size="25" value="<%=objNodes.childNodes(1).text%>"> </td>                                                                         
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;颜色：<input type="text" name="color" size="25" value="<%=objNodes.childNodes(3).text%>"> 输入颜色代码例如：#000000</td>      
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;头像：<br>
        <%dim fcount:for fcount=1 to 100%>
        <input type="radio" name="face" value="<%=fcount%>"<%if objNodes.childNodes(2).text=cstr(fcount) then response.write" checked"%>><img src="images/qqface/<%=fcount%>_m.gif" border="0">                                          
        <%next%>                                          
        </td>                                                                           
      </tr>
      <tr>
        <td width="500" height="10"></td>
      </tr>
      <tr>
        <td width="500" height="25">&nbsp;<input type="submit" name="submit" value="确定修改">&nbsp;<input type="reset" name="reset" value="取消重置"></td>
      </tr>
      <tr>
        <td width="500" height="10"></td>
      </tr>
    </table>
    </form>
    </fieldset>
    </td>
  </tr>
</table>
<%
else
%>
<script language="javascript">
alert("发生错误！")
history.back()
</script>
<%
end if
Set objNodes=nothing
Set objXML=nothing
end sub

if request.querystring("type")="manage" then
%>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500" height="20"></td>
  </tr>
</table>
<%
call cklogin()
select case request.querystring("mtype")
case "edit" call neditinfo()
case else call slist()
end select
dim enddtime:enddtime=timer()
%>
<table border="0" width="500" cellpadding="0" cellspacing="0" align="center">
  <tr>
    <td width="500" height="10"></td>
  </tr>
  <tr>
    <td width="500" height="20" align="center">通用网站在线咨询JQQonline插件 <a href="http://www.happyjh.com" target="_blank">happyjh.Com</a> 执行时间:<%=FormatNumber((enddtime-starttime)*1000,3)%>毫秒</td>
  </tr>
</table>
<%
else
response.clear
response.redirect "/"
response.end
end if
%>
</body>
</html>