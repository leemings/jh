<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>新增房间</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"><font color="#000000">房间新增程序</div>
<form method="post" action="addroomok.asp"><center>
  <table border="0" cellspacing="1" cellpadding="4" bgcolor="f2f2ea" cellspacing=0 cellpadding=0  width="75%">
    <tr> 
      <td width="67"> 
        <div align="center">操作</div>
      </td>
      <td width="197">新增江湖房间</td>
      <td width="81"> 
        <div align="center">房间名</div>
      </td>
      <td width="224">
        <input type="text" name="name" size="15" maxlength="20" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">最大人数</div>
      </td>
      <td width="197">
        <input type="text" name="zdrs" size="15" maxlength="3" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      <td width="81" nowrap> 
        <div align="center">发招限制</div>
      </td>
      <td width="224">
        <select name=fzxz size=1 >
          <option value='0' selected >允许</option>
          <option value='1' >禁止</option>
        </select>
        </td>
    </tr>
 <tr> 
      <td width="67"> 
        <div align="center">打架时间</div>
      </td>
      <td width="197"> 
        <input type="text" name="djsj"
value="60" size="2" maxlength="2" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid"> [设置60为不限制]
        </td>
      <td width="81" nowrap> 
        <div align="center">泡点</div>
      </td>
      <td width="224"> 
        <select name=pd size=1 >
          <option value='1' selected>允许</option>
          <option value='0'>禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center">事件限制</div>
      </td>
      <td width="197">
        <select name=sjxz size=1 >
          <option value='0' selected >允许</option>
          <option value='1' >禁止</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">保护</div>
      </td>
      <td width="224">
        <select name=bh size=1 >
          <option value='0' selected>允许</option>
          <option value='1'>禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">卡片</div>
      </td>
      <td width="197">
        <select name=kp size=1 >
          <option value='0' selected >允许</option>
          <option value='1' >禁止</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">赌博</div>
      </td>
      <td width="224">
        <select name=db size=1 >
          <option value='0' selected>允许</option>
          <option value='1' >禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">限制</div>
      </td>
      <td colspan="3">
        <select name=xz size=1>
          <option value='0' selected >允许[加入此房间无限制]</option>
          <option value='1' >禁止[加入此房间有限制]</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">限制说明</div>
      </td>
      <td colspan="3">
        <input type="text" name="xzsm" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">表达式</div>
      </td>
      <td colspan="3">
        <input type="text" name="bds" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        <br>
        可以使用：and、or、not、+、-、=等表达式请对应条件连接！<br>
        对于编程不了解者不可随意修改！(ID为1主聊天不要加限制！) </td>
    </tr>
    <tr> 
      <td width="67" height="2"> 
        <div align="center">今日话题</div>
      </td>
      <td colspan="3" height="2">
        <input type="text" name="jrht"
value="话题" size="50" maxlength="150" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        (在进入江湖时是会显示的)</td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="确 定">
          <font color="#CCCCCC">-------  
          <input  onClick="javascript:window.document.location.href='roomlist.asp'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<div align="center">数据库更新程序因为时间有限，没有加入为空值时的检测，请大家更改时注意没有值的地方填0</div>