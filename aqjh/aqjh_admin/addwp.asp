<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
lb=LCase(trim(Request.QueryString("lb")))
if lb="" then lb="药品"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>添加物品</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font size="2" color="#000000">添加物品</font></p>
<form method="post" action="addwpok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="386">
    <tr> 
      <td colspan="2" height="18"><font color="#FFFF00" size="-1">提示:对于添加卡片，还要改chat/sjfunc/44.asp文件!</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">物品名</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="a" size="20"
value="<%=lb%>">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">图片</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="i"
value="无" size="10">
        (hcjs/jhjs/images)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">耐久度</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="j"
value="0" size="10">
        (只对装备有效果)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">特性</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="k"
value="无" size="8">
        (只对<font color="#FFFF00">右手</font>及<font color="#FFFF00">宠物</font>有效果，其它填写:<font color="#FFFF00">无</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">条件</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="l" size="20"
value="True">
        (任意购买填写:<b><font color="#FFFF00">True</font></b>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">类别</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="b"
value="<%=lb%>" size="10">
        (根据情况填写)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">内力</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="d"
value="0" size="10">
        (对于<font color="#FFFF00">药品</font>、<font color="#FFFF00">暗器</font>、<font color="#FFFF00">毒药</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">体力</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="e"
value="0" size="10">
        (对于<font color="#FFFF00">药品</font>、<font color="#FFFF00">暗器</font>、<font color="#FFFF00">毒药</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">攻击</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="f"
value="0" size="10">
        (对于所有装备有效)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">防御</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="g"
value="0" size="10">
        (对于所有装备有效)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">说明</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="c" size="40"
value="<%=lb%>">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">银两</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="h"
value="0" size="10">
        (<font color="#FFFF00"><b>卡片</b></font>时填写:0 隐藏此卡)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">金币</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="m"
value="0" size="10">
        (<font color="#FFFF00"><b>只对宠物有效</b></font>)</font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="确 定">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='wpadmin.asp?lb=<%=lb%>'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
