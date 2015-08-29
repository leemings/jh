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
fl=Trim(Request.QueryString("fl"))
if fl<>"右手" and fl<>"左手" and fl<>"盔甲" and fl<>"头盔" and fl<>"双脚" and fl<>"装饰" and fl<>"药品" and fl<>"毒药" and fl<>"卡片" then
	Response.Write "<script Language=Javascript>alert('分类查只能为：右手、左手、盔甲、头盔、双脚、装饰、药品、毒药、卡片，请重新选择！');window.close();</script>"
	Response.End
end if
if sjjh_grade<9 then
	Response.Write "<script Language=Javascript>alert('你想作什么滚！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM w WHERE c='" & fl & "' and i>0  order by b",conn
%>
<html>
<head>
<title>物品管理♀一线网络→wWw.51eline.com♀</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center"><font color=yellow><b><%=fl%>类物品</b></font>一览(装备的物品除外)<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div>
<br>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">拥有者</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">类型 </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">数量 </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">内力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">体力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">攻击 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">防御 </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">价钱</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">装备</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">方式</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("b")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("a")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("c")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("i")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("e")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("f")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("g")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("h")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("l")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("j")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">删除</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM s WHERE e='" & fl & "' order by b,c",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">拥有者</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">类型 </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">数量 </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">内力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">体力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">攻击 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">防御 </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">价钱</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">方式</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">操作</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("b")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("a")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("e")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("j")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("f")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("g")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("h")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("i")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("n")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("c")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">删除</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<tr align="center"> 
  <td nowrap width="45" height="16"> 
    <div align="center"></div>
  </td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td height="25"> 
      <div align="center">这里可以查看到对方的物品，删除就可以将此用户的物品删除！</div>
    </td>
  </tr>
</table>
</body>
</html>
