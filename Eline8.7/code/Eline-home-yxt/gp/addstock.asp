<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>E线江湖增加股票</title></head>
<body text="#000000" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align=center><font color=0000ff size=4>增发新股</font></p>
<form action=addstocknow.asp method=post>
<table border=1 width=80% align=center cellpadding="0" cellspacing="0">
<tr><td>股票名称</td><td><input type=text name='stock' size=14 maxlength=7 value=''></td></tr><tr><td>发行数量</td><td><input type=text value='' size=9 maxlength=9 name='num'></td></tr><tr><td>发行价格</td><td><input type=text name='price' value='' size=9 maxlength=9></td></tr><tr align=center><td colspan=2><input type=submit name=opt value=' 发 行 '> <input type=button onclick=javascript:history.back(); value=' 返 回 '></td></tr></table></form></body>