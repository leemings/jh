<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 配偶<>'无' and 配偶<>姓名 and 姓名='"&sjjh_name&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&sjjh_name&"]你还没有配偶呢，来作什么！');history.go(-1);</script>"
	response.end
end if
rs.close
%>
<html>
<head>
<title>自创夫妻合体技</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../setup.css">
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"> <font color="#FFFFFF" ><b><font color="#000000">夫 妻 合 体 技 设 置</font></b></font><font color="#000000"><br>
<br>
</font> </div>
<table cellpadding="0" cellspacing="1" border="1" align="center" width="45%" bgcolor="#00CCFF" bordercolor="#000000">
  <tr> 
    <td height="11" width="42"> 
      <div align="center">ID</div>
    </td>
    <td height="11" width="97"> 
      <div align="center"> <font color="#000000" size="2">招 式 名</font> </div>
    </td>
    <td height="11" width="93"> 
      <div align="center"> <font color="#000000" size="2">所 用 内 力</font> </div>
    </td>
    <td height="11" width="104"> 
      <div align="center"> <font color="#000000" size="2">操 作</font> </div>
    </td>
  </tr>
  <%
rs.open "SELECT * FROM t where b='" & sjjh_name & "' or c='" & sjjh_name & "'"
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='stuntok.asp?a=m'>
    <tr> 
      <td height="2" width="42"> 
        <div align="center"><font color="#000000" size="2"><%=rs("id")%></font></div>
      </td>
      <td height="2" width="97"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="wg" size="10" value="<%=rs("a")%>" maxlength="8">
          <input type="hidden" name="id" size="10" value="<%=rs("id")%>" maxlength="8">
          </font> </div>
      </td>
      <td height="2" width="93"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="nl" size="10" value="<%=rs("d")%>" maxlength="8">
          </font> </div>
      </td>
      <td height="2" width="104"> 
        <div align="center"> <font color="#000000" size="2">
<input type="submit" value="修改" name="submit">
          <a href="del.asp?id=<%=rs("id")%>">删除</a></font> </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<5 then
%>
  <form method=POST action='stuntok.asp?a=n'>
    <tr> 
      <td width="42" height="2"> 
        <div align="center"><font color="#000000" size="2">新建</font></div>
      </td>
      <td width="97" height="2"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="wg" size="10" maxlength="8">
          </font> </div>
      </td>
      <td width="93" height="2"> 
        <div align="center"> <font color="#000000" size="2"> 
          <input type="text" name="nl" size="10" maxlength="8">
          </font> </div>
      </td>
      <td width="104" height="2"> 
        <div align="center"> <font color="#000000" size="2">
<input type="submit" value="添加" name="submit">
          </font> </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center"><font color="#000000" size="2">[注：只有已婚大侠或侠女才可以自创，每创建一次收取10000两，创建后双方都可使用，离婚后合体技失效！]<br>
<br>
  [杀伤力=<b><font color="#0000FF">((男方攻击+女方攻击)</font></b><font color="#FF0000">-对方防御</font>)+所用内力]<br>
<br>
<br>
【授权:『E线江湖』 运行序列号：<%=Application("sjjh_sn")%> 】
</font></p>
</body>
</html>