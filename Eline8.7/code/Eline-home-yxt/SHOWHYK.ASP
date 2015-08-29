<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
ID=request("ID")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_hykmdb")
rs.open "SELECT * FROM 会员卡 Where 卡号='"&ID&"'",conn
if rs.eof or rs.bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：找不到此卡！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
%>
<html>
<head>
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><font color="#0000FF">会员卡管理</font></p>
<p align="center"><font color="#0000FF">&nbsp;说明：如果密码不想更改的话请保持密码为“不更改字样”</font></p>
<p align="center"><font color="#0000FF">点数：点数按实际价格*1000计算，也就是说1RMG=1000点</font></p>
<p align="center"><font color="#0000FF">
类型：0为未使用，1为使用过，2冲为泡点会员分钟卡，3为冲为等级会员分钟卡，4为废卡，在添加新卡的时候</font></p>
<p align="center"><font color="#0000FF">
会自动覆盖。此项如果更改可能会出现不可预料的后果，所以禁止改动。</font></p>
<p align="center"><font color="#0000FF">时间为：在使用分钟会员的时候所剩的时间是多少如果类型不是2或3，改动是无效的~</font></p>
<form method="POST" action="updatehyk.asp?ID=<%=ID%>">
  <table border="1" width="27%" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="80" align="center">
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27">
    卡&nbsp; 号</td> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27"> 
    <input type="text" name="KNAME" size="16" value="<%=ID%>" maxlength="8"></td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27"> 
    密&nbsp; 码</td> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27"> 
    <input type="text" name="KPASS" size="16" value="不更改" maxlength="8"></td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27"> 
    点&nbsp; 数</td> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27"> 
    <input type="text" name="KNUM" size="16" value="<%=rs("点数")%>"></td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27"> 
    类&nbsp; 型</td> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27"> 
  <% 
    select case rs("类型") 
    case 0 
    %> 
    未使用 
    <% 
   case 1 
    %> 
    已使用 
    <% 
   case 2 
    %> 
   冲泡点分钟 
    <% 
   case 3 
    %>  
   冲等级分钟 
    <% 
   case 4 
    %>  
   已作废 
    <% 
end select 
    %> 
</td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27"> 
    用户ID</td> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27"> 
    <%=rs("用户ID")%></td> 
  </tr> 
  <tr>  
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27"> 
    时&nbsp; 间</td>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27">
    <input type="text" name="KTIME" size="16" value="<%=rs("时间")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="33%" align="center" nowrap height="27">
    <input type="submit" value="更新" name="submit"></td>
    <td bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#DFEFFF';"  width="68%" align="center" nowrap height="27">
    <input type="submit" value="删除" name="submit"></td>
  </tr>
  </table>
</form>
<p align="center">　</p>
<p align="center">本程序由阿男编写</p>
</body>
</html>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
%>