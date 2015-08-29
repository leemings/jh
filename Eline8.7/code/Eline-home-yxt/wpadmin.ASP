<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
lb=LCase(trim(Request.QueryString("lb")))
if lb="" then lb="药品"
%>
<html>

<head>
<title>『E线江湖』物品管理程序♀wWw.51eline.com♀</title>
<style></style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center">『E线江湖』物品管理程序<br>
  <br>
  <a href="wpadmin.asp?lb=左手">左手</a> <a href="wpadmin.asp?lb=右手">右手</a> 
  <a href="wpadmin.asp?lb=装饰">装饰</a> <a href="wpadmin.asp?lb=盔甲">盔甲</a> 
  <a href="wpadmin.asp?lb=双脚">双脚</a> <a href="wpadmin.asp?lb=头盔">头盔</a> 
  <a href="wpadmin.asp?lb=暗器">暗器</a> <a href="wpadmin.asp?lb=宠物">宠物</a> 
  <a href="wpadmin.asp?lb=药品">药品</a> <a href="wpadmin.asp?lb=毒药">毒药</a> 
  <a href="wpadmin.asp?lb=卡片">卡片</a> <a href="wpadmin.asp?lb=鲜花">鲜花</a> 
  <a href="wpadmin.asp?lb=锻造">锻造</a>
</p>
<table width="600" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000"
bgcolor="#006699">
  <tr> 
    <td colspan="4" height="11" bgcolor="#336633"> 
      <p align="center"><font color="#FFFF00">现 有 <b>[<%=lb%>]</b> 列 表</font></p>
    </td>
    <td height="11" bgcolor="#336633" width="84"> 
      <div align="center"><a href="addwp.asp?lb=<%=lb%>">增加物品</a></div>
    </td>
  </tr>
  <tr> 
    <td width="82" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">物品名</font></div>
    </td>
    <td width="63" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">类 别</font></div>
    </td>
    <td width="259" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">功 能</font> </div>
    </td>
    <td width="94" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">售 价</font> </div>
    </td>
    <td width="84" height="20" bgcolor="#C4DEFF"> 
      <div align="center"><font size="2" color="#000000">操 作</font> </div>
    </td>
  </tr>
  <%
sql="SELECT * FROM b where b='"&lb&"' order by b,h "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#85C2E0';"> 
    <td width="82" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("a")%> </font></div>
    </td>
    <td width="63" height="23"> 
      <div align="center"><font color="#FFFFFF" size="2"><%=rs("b")%></font></div>
    </td>
    <td width="430" height="23"> 
      <div align="left"><font color="#FFFFFF" size="2">
      <%select case lb
      	case "卡片","宠物","鲜花"
	      	Response.Write "说明:"&rs("c")
			if lb="宠物" then
				Response.Write " 特性:"&rs("k")
			end if
	      	if lb="宠物" or lb="鲜花" then
   		      	Response.Write " 图:<img src='../hcjs/jhjs/images/"&rs("i")&"'>"
	      	end if
		case "毒药","暗器","药品"
			Response.Write "内力："&abs(rs("d"))&" 体力："&abs(rs("e"))
		case else
			Response.Write  "攻击："&abs(rs("f"))&" 防御:"&abs(rs("g"))
			if lb="右手" or lb="锻造" then
			Response.Write " 特性:"&rs("k")
			end if
	      	if lb="左手" or lb="右手" or lb="盔甲" or lb="双脚" or lb="装饰" or lb="头盔" or lb="锻造" then
   		      	Response.Write " 耐久:"&rs("j")&" 图:<img src='../hcjs/jhjs/images/"&rs("i")&"'>"
	      	end if
      	end select
      %>
        
        </font></div>
    </td>
    <td width="94" height="23"> 
      <div align="right"><font color="#FFFFFF" size="2"><%=rs("h")%><%if lb<>"卡片" then%> 两<%else%> 元<%end if%></font></div>
    </td>
    <td width="84" height="23"> 
      <div align="center"><a
href="modiwp.asp?wupinid=<%=rs("id")%>"><font color="#FFFFFF" size="-1">修改</font></a> 
        <font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">删除</font></a></font></div>
    </td>
  </tr>
  <%
rs.movenext
loop
conn.close
%>
</table>
</body>
</html>
