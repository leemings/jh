<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>会员数据库管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif">
<p align="center"><%=Application("sjjh_chatroomname")%>数据库</p>
<table width="572" border="1" align="center" cellspacing="0" cellpadding="0" bordercolor="#000000">
<tr>
<td>
      <p align="center"><a href="hygl.asp">会员在线申请管理</a></p>
<form method="POST" action="hydjok.asp">
        <div align="center">
          <p><b><font color="#FF0000">注意：</font></b><font color="#FF0000">对于已经是会员的，设置完成后时间会按会员结束时间加<br>
            上设置时间，会员等级会设置成新的等级!</font></p>
          <p><font color="#FF0000"><br>
            </font>会员申请时间： 
            <select name=hymonth   size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
              <option value="1" selected>01</option>
              <option value="2">02</option>
              <option value="3">03</option>
              <option value="4">04</option>
              <option value="5">05</option>
              <option value="6">06</option>
              <option value="7">07</option>
              <option value="8">08</option>
              <option value="9">09</option>
              <option value="10">10</option>
              <option value="11">11</option>
              <option value="12">12</option>
            </select>
            月 <br>
            会员等级： 
            <select name="hygrade">
              <option value="1" selected>一级会员</option>
              <option value="2">二级会员</option>
              <option value="3">三级会员</option>
              <option value="4">四级会员</option>
            </select>
            <br>
            会员姓名： 
            <input type="text" name="search" size="10" maxlength="10">
            <br>
            <input type="submit" value="确定新会员" name="B1" class="p9">
          </p>
        </div>
      </form>
<div align="center"><font color="#FF0000"><br>
          千万不要拿站长设置会员，付费制会员请手动设置！</font><br>
<br>
</div>
<div align="center">请选择所要查姓名及等级设置！<br>
需要注意：会员截止日期不能小于当前日期，会员在到期前5天提示！<br>
不同等级的会员加的也不同的：1级会员：战斗25 2级会员：战斗40 上限加：100<br>
3级会员:战斗70级 上限加：200 4级：战斗98 管理:7级 上限：300 如果要修改请手动改！<br>
</div>
<p align="center">[站长可以通过此处来修改用户的所以资料]</p>
<p align="center"><a href="../jhmp/hy.asp?fs=1">现有会员列表</a> <a href="fine.asp">更新用户资料</a></p>
</td>
</tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>
