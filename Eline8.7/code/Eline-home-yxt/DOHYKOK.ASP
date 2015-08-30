<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.Buffer=true
name=Trim(Request.Form("name"))
pass=Trim(Request.Form("pass"))
kname=Trim(Request.Form("kname"))
kpass=Trim(Request.Form("kpass"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
pass=md5(pass)
kpass=md5(kpass)
sql="SELECT * FROM 用户 WHERE 姓名='"&name&"'"
rs.open sql,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('没有这个用户！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("密码")<>pass then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('用户密码不正确！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
userid=rs("id")
hydj=rs("会员等级")
hy=rs("会员")
rs.close
conn.close
set rs=nothing
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_hykmdb")
rs.open "SELECT * FROM 会员卡 WHERE 卡号='"&kname&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('这个卡号不存在！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("密码")<>kpass then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('会员卡密码不正确！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("类型")=4 or rs("类型")=2 or rs("类型")=3  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('这个卡号不存在！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("类型")=0 then
rs("类型")=1
rs.Update
end if
knum=rs("点数")
rs.close
rs.open "SELECT * FROM sethyk WHERE name='set'",conn
%>
<html>
<head>
<title>会员数据库管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">欢迎[<%=name%>]使用会员卡系统</font></p>
<p align="center"><font color="#0000FF">&nbsp;</font></p>
<p align="center">　</p>
<form method="POST" action="updohyk.asp?kname=<%=kname%>&id=<%=userid%>" name="updohyk">
  <table border="1" width="64%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="127" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="37" colspan="4">
    <p>您的卡里边还有[<%=knum%>]点！</td>
  </tr>
  <%
  if rs("dhyon") then
  rs.close
  %>
  <tr> 
    <td bgcolor="#85C2E0"   width="33%" align="center" nowrap height="37">
    等级制会员</td>
    <td bgcolor="#85C2E0"   width="4%" align="center" nowrap height="37">
    <select size="1" name="D1">
    <%if hydj=0 or hydj=1 then%>
    <option value="1">一级</option>
    <%end if%>
    <%if hydj=0 or hydj=2 then%>
    <option value="2">二级</option>
    <%end if%>
    <%if hydj=0 or hydj=3 then%>
    <option value="3">三级</option>
    <%end if%>
    <%if hydj=0 or hydj=4 then%>
    <option value="4">四级</option>
    <%end if%>
    </select></td>
    <%
    rs.open "SELECT * FROM 会员卡 Where 用户ID="&userid&" and 类型=3",conn
    %>
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="37">
    <%if rs.Eof and rs.Bof then%>
    <!--webbot bot="Validation" s-data-type="Integer" s-number-separators="x"
    --><input type="text" name="T1" size="4" value="0">天&nbsp;&nbsp;&nbsp;&nbsp; 天：<input type="radio" value="V1" name="R1">    
    <%end if%>   
    <%if (not rs.Eof and not rs.Bof) or hydj=0 then%>   
    分：<input type="radio" value="V2" name="R1">   
   <%   
    end if 
    rs.close 
    %> 
    </td> 
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="37"> 
     
    选定：<input type="radio" name="R3" value="V1">&nbsp; </td> 
  </tr> 
   
  <% 
  rs.open "SELECT * FROM sethyk WHERE name='set'",conn 
  end if 
  if rs("phyon") then 
  rs.close 
  %> 
  <tr>  
    <td bgcolor="#85C2E0"   width="33%" align="center" nowrap height="37"> 
    泡点制会员</td> 
    <td bgcolor="#85C2E0"   width="15%" align="center" nowrap height="37"> 
　</td> 
   <% 
    rs.open "SELECT * FROM 会员卡 Where 用户ID="&userid&" and 类型=2",conn 
    %> 
    <td bgcolor="#85C2E0"   width="40%" align="center" nowrap height="37"> 
      <%if rs.Eof and rs.Bof then%> 
      <!--webbot bot="Validation" s-data-type="Integer" s-number-separators="x"
      --><input type="text" name="T2" size="4" value="0">天&nbsp;&nbsp;&nbsp;&nbsp; 天：<input type="radio" value="V1" name="R2" checked>    
    <%   
    end if 
    if (not rs.Eof and not rs.Bof) or not hy then 
    %>&nbsp;   
    分：<input type="radio" value="V2" name="R2"> 
    <%   
    end if 
    rs.close 
    %> 
   </td> 
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="37"> 
    选定：<input type="radio" name="R3" value="V2">&nbsp; </td> 
  </tr> 
   <% 
  rs.open "SELECT * FROM sethyk WHERE name='set'",conn 
  end if 
  if rs("mon") then 
  rs.close 
  %> 
  <tr>  
    <td bgcolor="#85C2E0"   width="33%" align="center" nowrap height="37"> 
    江湖银子</td> 
    <td bgcolor="#85C2E0"   width="55%" align="center" nowrap height="37" colspan="2"> 
    <!--webbot bot="Validation" s-display-name="江湖银子必须大于0"
    s-data-type="Integer" s-number-separators="x"
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="T3" size="16" value="0">两</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="37">
    选定：<input type="radio" name="R3" value="V3">&nbsp; </td>
 <%
  rs.open "SELECT * FROM sethyk WHERE name='set'",conn
  end if
  if rs("jon") then
  rs.close
  %>
  <tr> 
    <td bgcolor="#85C2E0"   width="33%" align="center" nowrap height="37">
    会员金卡</td>
    <td bgcolor="#85C2E0"   width="55%" align="center" nowrap height="37" colspan="2">
    <!--webbot bot="Validation" s-display-name="会员金卡必须大于0"
    s-data-type="Integer" s-number-separators="x"
    s-validation-constraint="Greater than or equal to" s-validation-value="0"
    --><input type="text" name="T4" size="16" value="0">两</td>
    <td bgcolor="#85C2E0"   width="11%" align="center" nowrap height="37">
    选定：<input type="radio" name="R3" value="V4">&nbsp; </td>
 </tr>
 <%
  rs.open "SELECT * FROM sethyk WHERE name='set'",conn
  end if
%>
  <tr> 
 
    <td bgcolor="#85C2E0"   width="99%" align="center" nowrap height="37" colspan="4">
    <input type="submit" value="确   定" name="B1"></td>

 </tr>
  </table>
</form>
<table border="1" width="64%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="135" align="center">
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="28" colspan="5">
    说明：</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="70" colspan="5">
    如果现在你已经是等级制会员，你的唯一的选择就是现在<br>
    的会员等级。除非你等的会员到期。<p>如果你冲的是分钟的话，就会将卡中所有的点数冲净。</td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0"   width="51%" align="center" nowrap height="28" colspan="5">
    比率：</td>
  </tr>
  <tr>
    <td bgcolor="#66FFFF"   width="11%" align="center" nowrap height="28">
    等级制会员</td>
    <td bgcolor="#3399FF"   width="10%" align="center" nowrap height="28">
    一级</td>
    <td bgcolor="#3399FF"   width="10%" align="center" nowrap height="28">
    二级</td>
    <td bgcolor="#3399FF"   width="10%" align="center" nowrap height="28">
    三级</td>
    <td bgcolor="#3399FF"   width="10%" align="center" nowrap height="28">
    四级</td>
  </tr>
  <tr> 
    <td bgcolor="#CCCCFF"   width="11%" align="center" nowrap height="28">
    100分钟</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy1f")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy2f")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy3f")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy4f")%>点</td>
  </tr>
  <tr> 
    <td bgcolor="#CCCCFF"   width="11%" align="center" nowrap height="28">
    一天</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy1t")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy2t")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy3t")%>点</td>
    <td bgcolor="#9999FF"   width="10%" align="center" nowrap height="28">
    <%=rs("dhy4t")%>点</td>
  </tr>
  <tr> 
    <td bgcolor="#99FFCC"   width="11%" align="center" nowrap height="28">
    泡点制会员</td>
    <td bgcolor="#FFCCCC"   width="10%" align="center" nowrap height="28">
    100分钟</td>
    <td bgcolor="#FFCCCC"   width="10%" align="center" nowrap height="28">
    <%=rs("phyf")%>点</td>
    <td bgcolor="#FF99FF"   width="10%" align="center" nowrap height="28">
    一天</td>
    <td bgcolor="#FF99FF"   width="10%" align="center" nowrap height="28">
    <%=rs("phyt")%>点</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFF99"   width="11%" align="center" nowrap height="28">
    江湖银子</td>
    <td bgcolor="#FFCC00"   width="10%" align="center" nowrap height="28">
    1000W两</td>
    <td bgcolor="#FFCC00"   width="10%" align="center" nowrap height="28">
    <%=rs("m")%>点</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="28">
    　</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="28">
    　</td>
  </tr>
   <tr> 
    <td bgcolor="#CC9900"   width="11%" align="center" nowrap height="28">
    会员金卡</td>
    <td bgcolor="#FF9933"   width="10%" align="center" nowrap height="28">
    1元</td>
    <td bgcolor="#FF9933"   width="10%" align="center" nowrap height="28">
    <%=rs("j")%>点</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="28">
    　</td>
    <td bgcolor="#85C2E0"   width="10%" align="center" nowrap height="28">
    　</td>
  </tr>

  </table>
<p align="center">『快乐江湖』</p>
</body>
</html>
<%
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>