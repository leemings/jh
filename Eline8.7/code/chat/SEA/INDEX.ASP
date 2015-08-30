<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
action=Request.Querystring("action")
if action="search" then
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
Randomize timer
r=int(rnd*10)
Set Conn1=Server.CreateObject("ADODB.Connection")
Connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Eline_global.asa")
Conn1.Open Connstr1
set rs=conn1.execute("select 次数,时间 from 航海物品 where 拥有者='"& name &"'")
if rs.eof and rs.bof then
  Conn1.Execute("insert into 航海物品(拥有者) values('"& name &"')")
times=8
else
times=rs("次数")
  if now()-rs("时间")>1 then
    conn1.execute("update 航海物品 set 次数=次数+8,时间=now() where 拥有者='"& name &"'")
    times=times+8
  end if
end if
        if times<1 then
            mess="你的行动力没了！！！"
        else
            conn1.execute("update 航海物品 set 次数=次数-1 where 拥有者='"& name &"'")
            if r=5 then
              conn.execute("update 用户 set 银两=银两+100000 where 姓名='"& name &"'")
              mess="发现了宝藏10万两！！！"
            else
              mess="一无所获！！！"
            end if
        end if
set rs=nothing
conn1.close : set conn1=nothing
conn.close : set conn=nothing
response.write "<script language='javascript'>alert('"& mess &"');window.close();</script>"
else
%>
<html>
<head>
<title>航海时代♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 12px}
input {  font-family: "宋体"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong>江湖航海时代</strong></big><br>
      <br>
      (每天可做<font face="Verdana">8</font>笔傻瓜似的生意) 
      <input type="button" Onclick="MM_openBrWindow('see.asp','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="商品一缆"> 
      此插件测试中，发现错误请告诉我！！！ <br> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000">  
    <td width="25%"><font color="#FFFFFF">大连港</font></td> 
    <td width="25%"><font color="#FFFFFF">杭州港</font></td> 
    <td width="25%"><font color="#FFFFFF">宁波港</font></td> 
    <td width="25%"><font color="#FFFFFF">泉洲港</font></td> 
  </tr> 
  <tr align="center">  
    <td width="25%">  
      <input type="button" Onclick="MM_openBrWindow('jys.asp?port=大连','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所"> 
      <input type="button" Onclick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城门古迹"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=杭州','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城门古迹" name="button"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=宁波','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button2"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城门古迹" name="button2"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=泉洲','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button3"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城门古迹" name="button3"> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center"> 
  <tr align="center">  
    <td height="22"><br>『快乐江湖』<br>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>
