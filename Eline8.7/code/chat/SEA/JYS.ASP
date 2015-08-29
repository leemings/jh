<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=sjjh_name
action=Request("action")
if action="true" then
if instr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("SERVER_NAME"))<>0 then
port=Request.Form("port")
what=Request.Form("type")
submit=trim(Request.Form("submit"))
num=trim(Request.Form("num"))
On Error Resume Next
num=Clng(num)
if err.number>0 then
  err.clear
  response.write "<script language='javascript'>alert('请输入数字！！！');history.back(-1);</script>"
  response.end
end if
if num<1 or num=empty then num=1
if num>10 then num=10
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Eline_global.asa")
Conn.Open Connstr
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open Application("sjjh_usermdb")
Set RS=Conn.Execute("select 拥有者,次数,时间 from 航海物品 where 拥有者='"& name &"'")
if RS.eof and RS.bof then
  Conn.Execute("insert into 航海物品(拥有者) values('"& name &"')")
  times=8
else
times=rs("次数")
  if now()-rs("时间")>1 then
    conn.execute("update 航海物品 set 次数=次数+800,时间=now() where 拥有者='"& name &"'")
    times=times+800
  end if
end if
RS.close
if times<1 then
mess="你的行动力没了！！！"
else
Rs=Conn1.execute("select 银两 from 用户 where 姓名='"& name &"'")
yin=Rs("银两")
Sql="select * from 航海时代 where 港口='"& port &"'"
Rs=Conn.Execute(Sql)
jia1=Rs("丝绸") : jia2=Rs("人参") : jia3=Rs("珠宝") : jia4=Rs("瓷器")
  if submit="买" then
    Select Case what
      Case "丝绸"
          if num*jia1>yin then
            mess="你的钱不够了！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两-"& num*jia1 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,丝绸=丝绸+"& num &" where 拥有者='"& name &"'")
            Conn.Execute("update 航海时代 set 丝绸=丝绸+1 where 港口='"& port &"'")
            mess="成功购买！！！"
          end if
      Case "人参"
          if num*jia2>yin then
            mess="你的钱不够了！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两-"& num*jia2 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,人参=人参+"& num &" where 拥有者='"& name &"'")
            Conn.Execute("update 航海时代 set 人参=人参+2 where 港口='"& port &"'")
            mess="成功购买！！！"
          end if
      Case "珠宝"
          if num*jia3>yin then
            mess="你的钱不够了！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两-"& num*jia3 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,珠宝=珠宝+"& num &" where 拥有者='"& name &"'")
            Conn.Execute("update 航海时代 set 珠宝=珠宝+3 where 港口='"& port &"'")
            mess="成功购买！！！"
          end if
      Case "瓷器"
          if num*jia4>yin then
            mess="你的钱不够了！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两-"& num*jia4 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,瓷器=瓷器+"& num &" where 拥有者='"& name &"'")
            Conn.Execute("update 航海时代 set 瓷器=瓷器+1 where 港口='"& port &"'")
            mess="成功购买！！！"
          end if
     End Select
  else
    Select Case what
      Case "丝绸"
          rs=Conn.Execute("select 丝绸 from 航海物品 where 拥有者='"& name &"'")
          if rs("丝绸")<num then
            mess="你没那么多丝绸！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两+"& num*jia1 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,丝绸=丝绸-"& num &" where 拥有者='"& name &"'")
            if jia1>1 then
               Conn.Execute("update 航海时代 set 丝绸=丝绸-1 where 港口='"& port &"'")
            end if
            mess="成功卖出！！！"
          end if
      Case "人参"
          rs=Conn.Execute("select 人参 from 航海物品 where 拥有者='"& name &"'")
          if rs("人参")<num then
            mess="你没那么多人参！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两+"& num*jia2 &" where 姓名='"& my &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,人参=人参-"& num &" where 拥有者='"& name &"'")
            if jia1>2 then
               Conn.Execute("update 航海时代 set 人参=人参-2 where 港口='"& port &"'")
            end if     
            mess="成功卖出！！！"
          end if
      Case "珠宝"
          rs=Conn.Execute("select 珠宝 from 航海物品 where 拥有者='"& name &"'")
          if rs("珠宝")<num then
            mess="你没那么多珠宝！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两+"& num*jia3 &" where 姓名='"& my &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,珠宝=珠宝-"& num &" where 拥有者='"& name &"'")
            if jia1>3 then
               Conn.Execute("update 航海时代 set 珠宝=珠宝-3 where 港口='"& port &"'")
            end if
            mess="成功卖出！！！"
          end if
      Case "瓷器"
          rs=Conn.Execute("select 瓷器 from 航海物品 where 拥有者='"& name &"'")
          if rs("瓷器")<num then
            mess="你没那么多瓷器！！！"
          else
            Conn1.Execute("update 用户 set 银两=银两+"& num*jia4 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次数=次数-1,瓷器=瓷器-"& num &" where 拥有者='"& name &"'")
            if jia1>1 then
               Conn.Execute("update 航海时代 set 瓷器=瓷器-1 where 港口='"& port &"'")
            end if
            mess="成功卖出！！！"
          end if
     End Select
  end if
end if
set Rs=nothing
Conn.close : set Conn=nothing
Conn1.close : set Conn1=nothing
response.write "<script language='javascript'>alert('"& mess &"');history.back(-1);</script>"
else
response.write "请不要作弊！！！"
end if
else
port=Request.Querystring("port")
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Eline_global.asa")
Conn.Open Connstr
Sql="select * from 航海时代 where 港口='"& port &"'"
Set Rs=Conn.Execute(Sql)
%>
<html>
<head>
<title><%=port%>港商会</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
input {  font-family: "宋体"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<table width="90%" border="1" cellspacing="2" cellpadding="2" align="center">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td colspan="4" height="22"><%=port%>港商会 <br>
      <br>
      (一次至多购买10单位)<br>
    </td>
  </tr>
  <tr align="center" bgcolor="#660000"> 
    <td height="20"><font color="#FFFFFF">丝绸</font></td>
    <td height="20"><font color="#FFFFFF">人参</font></td>
    <td height="20"><font color="#FFFFFF">珠宝</font></td>
    <td height="20"><font color="#FFFFFF">瓷器</font></td>
    <%do while not Rs.eof%> </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="20"><%=Rs("丝绸")%>两/单位</td>
    <td height="20"><%=Rs("人参")%>两/单位</td>
    <td height="20"><%=Rs("珠宝")%>两/单位</td>
    <td height="20"><%=Rs("瓷器")%>两/单位</td>
  </tr>
  <tr align="center">
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 买 " OnClick=check()>
        <input type="submit" name="submit" value=" 卖 " OnClick=check()>
        <input type="hidden" name="type" value="丝绸">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 买 " OnClick=check()>
        <input type="submit" name="submit" value=" 卖 " OnClick=check()>
        <input type="hidden" name="type" value="人参">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 买 " OnClick=check()>
        <input type="submit" name="submit" value=" 卖 " OnClick=check()>
        <input type="hidden" name="type" value="珠宝">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 买 " OnClick=check()>
        <input type="submit" name="submit" value=" 卖 " OnClick=check()>
        <input type="hidden" name="type" value="瓷器">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
  </tr>
  <%
Rs.movenext
Loop
Rs.close : set Rs=nothing
Conn.close : set Conn=nothing
end if
%> 
</table>
</body>
</html>