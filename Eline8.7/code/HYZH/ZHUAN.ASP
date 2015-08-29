<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>会员充值系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<body bgcolor="#FFFFFF" background="../chat/img/Bk_hc3w.gif">
<p align="center"><b><font color="#008000" size="5"><img border="0" src="../chat/img/bgdz.gif"></font><font size="5" color="#008080">会员转换系统<img border="0" src="../chat/img/yuanwu.gif"></font></b></p>
<p align="center"><a href="czk.asp"><font size="3">充值卡</font></a><font size="3"> 
<a href="../hc3w_Admin/DOHYK.ASP"><font color="#800000">点数卡</font></a></font></p> 
<table width="712" border="1" align="center" cellspacing="0" cellpadding="0" bordercolor="#FFFF00"> 
<tr> 
<td width="708" bgcolor="#008080" bordercolor="#0000FF"> 
<form method="POST" action="zhuan1.asp"> 
        <div align="center"> 
            <font color="#008080"> 
            <input type="submit" value="转换1级会员" name="B1" class="p9"> 
            </font> 
        </div> 
      </form> 
<form method="POST" action="zhuan2.asp"> 
        <div align="center"> 
            <font color="#008080"> 
            <input type="submit" value="转换2级会员" name="B1" class="p9"> 
            </font> 
        </div> 
      </form> 
<form method="POST" action="zhuan3.asp"> 
        <div align="center"> 
            <font color="#008080"> 
            <input type="submit" value="转换3级会员" name="B1" class="p9"> 
            </font> 
        </div> 
      </form> 
<form method="POST" action="zhuan4.asp"> 
        <div align="center"> 
            <font color="#008080"> 
            <input type="submit" value="转换4级会员" name="B1" class="p9"> 
            </font> 
        </div> 
      </form> 
<div align="center"><font color="#FFFFFF">可以使用自己的金卡、金币、银两、充值卡换取会员（切忌要满足条件）<br>一级会员/月需要100元金卡、100个金币、银两1亿、充值卡1张。加存点31250到期后减掉</font><font color="#FFFFFF">31250</font>

</div>

<div align="center"><font color="#FFFFFF">二级会员/月需要150元金卡、150个金币、银两2亿、充值卡2张。加存点90000到期后减掉90000<br>三级会员/月需要200元金卡、200个金币、银两3亿、充值卡3张。加存点250000到期后减掉250000<br>四级会员/月需要300元金卡、300个金币、银两4亿、充值卡4张。加存点490000到期后减掉</font><font color="#FFFFFF">490000</font><font color="#008080"><br>

  </font>

</div>

</tr>
</table>
</body>
</html>
