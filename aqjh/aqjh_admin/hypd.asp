<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title>泡点制会员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>数据库</p>
<table width="572" border="1" cellspacing="1" cellpadding="4" align="center" bordercolor="#B8AF86">
<tr>
    <td height="189"> 
<form method="POST" action="hypdok.asp">
        <div align="center">
          <p><b><font color="#FF0000">注意：</font></b><font color="#FF0000">对于已经是会员的，设置完成后时间会按会员结束时间加<br>
            上设置时间，会员等级会设置成新的等级!</font></p>
          <p><font color="#0000FF">说明：泡点成倍制会员，为6.2版新增功能！设置成会员后他在聊天中泡点<br>
            为正常的N倍 ，N可以在paodian.asp中设置！</font><font color="#FF0000"><br>
            (<font color="#0000FF"><b>泡点成倍制会员不分等级！</b></font>)<br>
            <br>
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
            会员姓名： 
            <input type="text" name="search" size="10" maxlength="10">
            <br>
            <input type="submit" value="确定新会员" name="B1" class="p9">
          </p>
        </div>
      </form>
      <div align="center"><font color="#FF0000"><br>
        </font></div>
</td>
</tr>
</table>
<p align="center"><a href="../jhmp/hy.asp?fs=1">现有会员列表</a> <a href="fine.asp">更新用户资料</a></p>
</body>
</html>