<%
response.buffer=true
session("stopp")=request("stopp")
session("nomoney")=request("nomoney")
session("addmoney")=request("addmoney")
session("myhouse")=request("myhouse")
session("pmoney")=request("pmoney")

if session("stopp")<>"" then money=money+400
if session("nomoney")<>"" then money=money+400
if session("addmoney")<>"" then money=money+1000
if session("myhouse")<>"" then money=money+1000
if session("pmoney")<>"" then money=money+10000
i=trim(cstr(session("players")))
application("money"&i)=50000-money
response.write i

if session("stopp")<>"" or session("nomoney")<>"" or session("addmoney")<>"" or  session("pmoney")<>"" then
response.redirect "rframe.asp"
end if
%>
<html>
<head>
<title>E线江湖--游戏在线</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>

<body bgcolor="#FFFFFF" background="img/bg.gif">
<form method="post" action="charp.asp">
  <table width="100%" border="0" cellspacing="0">
    <tr> 
      <td width="40%"> 
        <div align="right"><font color="#FFFF33">请设置玩家数目</font></div>
      </td>
      <td width="30%"> 
        <div align="center"> <font color="#FFFF33"> 
          <input type="radio" name="playall" value="2">
          二位 
          <input type="radio" name="playall" value="3">
          三位 
          <input type="radio" name="playall" value="4">
          四位 </font></div>
      </td>
      <td>　</td>
    </tr>
  </table>
  <table width="789" border="0" cellspacing="0">
    <tr> 
      <td width="8%">　</td>
      <td width="34%">　</td>
      <td width="42%">　</td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td colspan="2"> 
        <div align="center"><font color="#66FF66"><b>请选择你要购买的卡片(至少要购买一种卡片)</b></font></div>
      </td>
      <td>　</td>
    </tr>
    <tr>
      <td width="8%">　</td>
      <td width="34%"><font color="#FFFF33">现金;50,000</font></td>
      <td width="42%">　</td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">停留卡 
          <input type="checkbox" name="stopp" value="stopp" checked>
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">400元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">免税卡 
          <input type="checkbox" name="nomoney" value="nomoney" checked>
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">400元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">查税卡 
          <input type="checkbox" name="addmoney" value="addmoney" checked>
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">1000元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">购地卡 
          <input type="checkbox" name="myhouse" value="myhouse" checked>
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">1000元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">均富卡 
          <input type="checkbox" name="pmoney" value="pmoney">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">10,000元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%">　</td>
      <td width="42%">　</td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">停留卡 
          <input type="checkbox" name="stopp2" value="stopp">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">400元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">免税卡 
          <input type="checkbox" name="nomoney2" value="nomoney">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">400元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">查税卡 
          <input type="checkbox" name="addmoney2" value="addmoney">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">1000元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">购地卡 
          <input type="checkbox" name="myhouse2" value="myhouse">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">1000元</font></td>
      <td>　</td>
    </tr>
    <tr> 
      <td width="8%">　</td>
      <td width="34%"> 
        <div align="right"><font color="#FFFF33">均富卡 
          <input type="checkbox" name="pmoney2" value="pmoney">
          </font></div>
      </td>
      <td width="42%"><font color="#FFFF33">10,000元</font></td>
      <td>　</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0">
    <tr>
      <td width="5%">　</td>
      <td width="70%">　</td>
      <td width="25%">　</td>
    </tr>
    <tr> 
      <td width="5%">　</td>
      <td width="70%"> 
        <div align="center"> 
          <input type="submit" name="Submit" value="确定">
          <input type="reset" name="Submit2" value="取消">
        </div>
      </td>
      <td width="25%">　</td>
    </tr>
  </table>
</form>
<p align="center"><span style="background-color: #FFFFFF" lang="zh-cn">『E线江湖』</span></p>
</body>
</html>