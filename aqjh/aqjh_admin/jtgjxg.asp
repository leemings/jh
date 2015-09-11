<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 车 where id="& id,conn
l=rs("类型")
a1=""
a2=""
a3=""
a4=""
a5=""
a6=""
a7=""
a8=""
select case l
	case "飞机"
		a1="selected"
	case "轿车"
		a2="selected"
	case "摩托车"
		a3="selected"
	case "火车"
		a4="selected"
	case "自行车"
		a5="selected"
	case "动物"
		a6="selected"
end select
%>
<html>
<head>
<title>交通工具修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../JHIMG/bk_Hc3w.gif">
<p align="center">修改交通工具</p>
<form name="form1" method="post" action="jtgjxgok.asp?id=<%=rs("id")%>">
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    类型：    
    <select name="lx">
      <option value="飞机" <%=a1%>>飞机</option>
      <option value="轿车" <%=a2%>>轿车</option>
      <option value="摩托车" <%=a3%>>摩托车</option>
      <option value="火车" <%=a4%>>火车</option>
      <option value="自行车" <%=a5%>>自行车</option>
      <option value="动物" <%=a6%>>动物</option>
    </select>
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  图片：     
    <input type="text" name="tp" value="<%=rs("图片")%>">
  <font color="#FF0000">注：此处图片不要加路径</font></font></p>   
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  价格：     
    <input type="text" name="jg" value="<%=rs("价格")%>">
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
  战斗等级：     
    <input type="text" name="zddj" value="<%=rs("战斗等级")%>" size="20">
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  &nbsp;&nbsp;&nbsp; 管理等级： </font><font size="2"><input type="text" name="gldj" value="<%=rs("管理等级")%>" size="20">  
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  会员专用 ：     
    <%if rs("会员专用")=true then%>
    <select name="hyzy">    
      <option value="true" selected>是</option>
      <option value="false">否</option>
    </select>
    <%else%>    
    <select name="hyzy">    
      <option value="true" >是</option>
      <option value="false" selected>否</option>
    </select>
    <%end if%>    
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  说明：     
    <textarea name="sm" rows="4" cols="56"><%=rs("说明")%></textarea>
    </font></p>
  <p align="center"><font size="2">说明里的图片位置要加准确的路径</font></p>
  <p align="center"> <font size="2"> 
    <input type="submit" name="Submit" value="修改">
    <input type="reset" name="重写" value="重置">
    </font></p>
</form>
<font size="2"> 
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</font> 
<p align="center"><font size="2"><a href="manjtgj.asp?lx=<%=l%>">返 回</a> </font></p>
</body>
</html>
