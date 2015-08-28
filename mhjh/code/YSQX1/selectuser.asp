<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=20 background="../chatroom/bg1.gif">
<div align=center>
分类查询查询<form method="POST" action="showuser2.asp">
  <div align="center"> 
    <select name="sjcz">
      <option value="姓名" selected>姓名</option>
      <option value="ID">ID</option>
      <option value="配偶">配偶</option>
      <option value="电子邮箱">信箱</option>
      <option value="帐号">帐号</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="15" maxlength="20">     
    <input type="submit" value="查询" name="B1" class="p9">     
    <input type="reset" name="Submit" value="清空">     
  </div>     
  <div align="center">ID查找一定要为数字！<br>     
  </div>     
  </form>    
按IP查询<form action="showbyip.asp" method=post><input type=text name=ip maxlength=15 size=15 value=''><input type=submit value=' 查 找 '></form>   
按等级查询<form action=showbygr.asp method=post>   
<select name=grade>   
<option value=-5>-5级</option>
<option value=-4>-4级</option>
<option value=-3>-3级</option>
<option value=-2>-2级</option>
<option value=-1>-1级</option>
<option value=0>0级</option>
<option value=1 selected>1级</option>
<option value=2>2级</option>
<option value=3>3级</option>
<option value=4>4级</option>
<option value=5>5级</option>
<option value=6>6级</option>
<option value=7>7级</option>
<option value=8>8级</option>
<option value=9>9级</option>
<option value=10>10级</option>
<option value=11>11级</option>
<option value=12>12级</option>
<option value=13>13级</option>
<option value=14>14级</option>
<option value=15>15级</option>
<option value=16>16级</option>
<option value=17>17级</option>
<option value=18>18级</option>
<option value=19>19级</option>
<option value=20>20级</option>
<option value=21>21级</option>
<option value=21>21级</option>
<option value=23>23级</option>
<option value=24>24级</option>
<option value=25>25级</option>
<option value=26>26级</option>
<option value=27>27级</option>
<option value=28>28级</option>
<option value=29>29级</option>
<option value=30>30级</option>
<option value=31>31级</option>
<option value=32>32级</option>
<option value=33>33级</option>
<option value=34>34级</option>
<option value=35>35级</option>
<option value=36>36级</option>
<option value=37>37级</option>
<option value=34>38级</option>
<option value=35>39级</option>
<option value=36>40级</option>
<option value=37>41级</option>
</select><input type=submit value='查找'></form>
按门派查询<form action=showbyco.asp method=post><select name=corp>
<option value='无' selected>无门无派</option>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "门派",conn
do while not rst.EOF
Response.Write "<option value='"&rst("门派")&"'>"&rst("门派")&"</option>"
rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
</select><input type=submit value='查找'></form>
删除过期帐号<form action=deluser.asp method=post>
<select name=howday>
<option value=7>7天</option>
<option value=15>15天</option>
<option value=30 selected>30天</option>
<option value=60>60天</option>
<option value=90>90天</option>
</select>
<input type=submit value='确认删除'>   
</form>   
    
<form method="POST" action="laren.asp">     
  <div align="center">拉人查看程序！<br>     
    <input type="text" name="uname" size="10" maxlength="10">     
    <input type="submit" value="查询" name="B12" class="p9">     
    <input type="reset" name="Submit2" value="清空">      
  <div align="center"></div>     
</form>     
<form method="POST" action="password.asp">     
  <div align="center">密码修正程序<br>输入人名将他的密码改成：123456<br>     
    <input type="text" name="cpass" size="10" maxlength="10">     
    <input type="submit" value="修改" name="B12" class="p9">     
    <input type="reset" name="Submit2" value="清空">     
    <br>     
  </div>     
  </form>     
</body>





