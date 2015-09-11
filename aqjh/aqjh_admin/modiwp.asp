<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sqlstr="select * from b where id="&wupinid
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Redirect "wpadmin.asp"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<title>物品修改</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font color="#000000"><br>
  对于物品修改，需要注意不要乱修改，否则会引起数据库的出错！<br>
  建议在修改之前作备份！ </font></p>
<form method="post" action="modiwpok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="386">
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinid" readonly
value="<%=rs("ID")%>" size="5">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">物品名</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="a"
value="<%=rs("a")%>" size="20">
        <%if rs("i")<>"无" then%>
        <img src="../hcjs/jhjs/images/<%=rs("i")%>"> 
        <%end if%>
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">图片</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="i"
value="<%=rs("i")%>" size="10">
        (hcjs/jhjs/images)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">耐久度</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="j"
value="<%=rs("j")%>" size="10">
        (只对装备有效果)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">特性</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="k"
value="<%=rs("k")%>" size="8">
        (只对<font color="#FFFF00">右手</font>及<font color="#FFFF00">宠物</font>有效果，其它填写:<font color="#FFFF00">无</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">条件</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="l"
value="<%=rs("l")%>" size="20">
        (任意购买填写:<b><font color="#FFFF00">True</font></b>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">类型</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="b"
value="<%=rs("b")%>" size="10">
        (根据情况填写)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">内力</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="d"
value="<%=rs("d")%>" size="10">
        (对于<font color="#FFFF00">药品</font>、<font color="#FFFF00">暗器</font>、<font color="#FFFF00">毒药</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">体力</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="e"
value="<%=rs("e")%>" size="10">
        (对于<font color="#FFFF00">药品</font>、<font color="#FFFF00">暗器</font>、<font color="#FFFF00">毒药</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">攻击</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="f"
value="<%=rs("f")%>" size="10">
        (对于所有装备有效)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">防御</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="g"
value="<%=rs("g")%>" size="10">
        (对于所有装备有效)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">说明</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="c"
value="<%=rs("c")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">银两</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="h"
value="<%=rs("h")%>" size="10">
        (<font color="#FFFF00"><b>卡片</b></font>时填写:0 隐藏此卡)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">金币</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="m"
value="<%=rs("m")%>" size="10">
        (<font color="#FFFF00"><b>只对宠物有效</b></font>)</font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="确 定">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='wpadmin.asp?lb=<%=rs("b")%>'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<div align="center">数据库更新程序因为时间有限，没有加入为空值时的检测，请大家更改时注意没有值的地方填0</div>