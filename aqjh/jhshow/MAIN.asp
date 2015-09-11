<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
if aqjh_name="" then
Response.Redirect "login.asp"
Response.End
end if
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
	
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from use where a='"& aqjh_name &"'", conn, 1,3
if rs.eof and rs.bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "reg.asp"
Response.End
end if
rs("h")=now()
rs.update
par4=rs("c")
if par4="M" then
mysex="男"
else
mysex="女"
end if
myjb=rs("e")
Response.Write "<script Language=Javascript>"
if rs("d")=true then
Response.Write "var myhy=1;"
else
Response.Write "var myhy=0;"
end if
Response.Write "var myshow='"& rs("f") &"';"
Response.Write "var par4='"& par4 &"';"
Response.Write "</script>"
rs.close
id = Request("id")
sex = Request("sex")
jbmoney=Request("money")
if jbmoney="" or jbmoney=-1 then jbmoney=999999999
if sex<>"M" and sex<>"F" then sex=""
search=trim(server.HTMLEncode(Request.QueryString("search")))
sid=Request("sid")
if sid="" then sid=801
rs.Open "select * from fl where fltype="& sid , conn, 1,1
idsm=rs("flname")
rs.close
if search="" then
	idsql="select * from data where d<>'"& sex &"' and f<="& jbmoney &" and j='"& sid &"'"
else
	idsql="select * from data where d<>'"& sex &"' and f<="& jbmoney &" and j='"& sid &"' and (b like '%" & search & "%' or c like '%" & search & "%') order by k desc"
end if
Response.Write"<HTML><HEAD><TITLE>江湖秀</TITLE><META http-equiv=Content-Type content=text/html; charset=gb2312><LINK href=images/show.css rel=stylesheet></HEAD>"
Response.Write"<BODY text=#000000 leftMargin=0 background=images/bg.gif topMargin=0  marginheight=0 marginwidth=0>"
Response.Write"<table width=100% height=865 border=0 align=center cellpadding=0 cellspacing=0><tr><td width=21% height=852 align=left valign=top bgcolor=#FF9900>"
Response.Write"<table width=97% border=1 align=center cellpadding=0 cellspacing=0 bordercolor=#000000><tr><td height=444 align=center valign=top> <table width=140 border=0 align=center cellpadding=0 cellspacing=0 bordercolorlight=#00215a bordercolordark=#000000>"
Response.Write"<tr><td width=140 height=226 align=center valign=middle bgcolor=#FFFFFF >"
Response.Write"<div id=show style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0; HEIGHT: 226px;'></div>"
Response.Write"</td></tr><tr><td height=44 align=center bgcolor=#CCCC99> <A href=javascript:userAv();><IMG height=20 src=images/main_right_24.gif width=55 border=0></A>&nbsp;<A href=javascript:userPhoto();><IMG height=20 src=images/main_right_25.gif width=55 border=0></A><BR>"
Response.Write"<A onclick=javascript:initshow(); href=#><IMG height=20 src=images/main_right_26.gif width=55 border=0></A>&nbsp;<A onclick=javascipt:deleteshow(); href=main.asp?sid="& sid &"&sex="& sex &"&CurPage=1><IMG height=20 src=images/main_right_27.gif width=55 border=0></A></td>"
Response.Write"</tr></table><TABLE cellSpacing=2 cellPadding=0 width=140 border=0><TBODY><TR><TD width=38 height=19>姓名：</TD><TD>&nbsp;"& aqjh_name &"</TD></TR><TR><TD height=19>性别：</TD>"
Response.Write"<TD>&nbsp;"& mysex &" <SCRIPT>if(myhy==1){document.write('<img src=images/vip_gif.gif>');}else{document.write('<font color=red>普通用户</font>')}</SCRIPT></TD></TR><TR><TD width=38 height=19><IMG src=images/left_tu_10.gif></TD>"
Response.Write"<TD bgColor=#e6e7e9>&nbsp;金币:"& myjb &"个</TD></TR><TR><TD width=38 height=19><IMG src=images/left_tu_04.gif width=38 height=19></TD><TD bgColor=#e6e7e9><A href=hy.asp target=_blank>&nbsp;帐户充值</A></TD>"
Response.Write"</TR><TR><TD width=38 height=19><IMG src=images/left_tu_03.gif width=38 height=19></TD><TD bgColor=#e6e7e9><A href=mysp.asp>&nbsp;物品管理</A></TD></TR><TR><TD width=38 height=19><IMG src=images/left_tu_05.gif width=38 height=19></TD><TD bgColor=#e6e7e9><A href=# onClick=javascript:faceurl()>&nbsp;形像代码</A></TD>"
Response.Write"</TR><TR><TD width=38 height=19><IMG src=images/left_tu_07.gif width=38 height=19></TD><TD bgColor=#e6e7e9>&nbsp;<strike>购买记录</strike></TD></TR><TR><TD width=38 height=19><IMG height=19 src=images/left_tu_08.gif width=38></TD>"
Response.Write"<TD bgColor=#e6e7e9>&nbsp;<strike>赠送记录</strike></TD></TR><TR><TD width=38 height=19><IMG height=19 src=images/left_tu_09.gif width=38></TD><TD bgColor=#e6e7e9>&nbsp;<strike>获赠记录</strike></TD></TR>"
Response.Write"<TR><TD width=38 height=19><IMG height=19 src=images/left_tu_11.gif width=38></TD><TD bgColor=#e6e7e9>&nbsp;<a href=exit.asp>退出</a></TD></TR></TBODY></TABLE></td></tr></table>"
Response.Write"<br><TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR><TD><TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR><TD><IMG height=19 src=images/main_left_31.gif width=162></TD>"
Response.Write"</TR></TBODY></TABLE><TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR><TD width=9 background=images/main_left_32.gif>&nbsp;</TD><TD vAlign=top width=140 background=images/main_left_39.gif height=30> <P><IMG height=24 src=images/main_left_35.gif width=101></P></TD>"
Response.Write"<TD width=13 background=images/main_left_33.gif>&nbsp;</TD></TR></TBODY></TABLE><TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR>"
Response.Write"<TD width=9 height=108 background=images/main_left_32.gif>&nbsp;</TD><TD vAlign=top width=140 background=images/main_left_39.gif><script language=javascript src=top.asp></script></TD>"
Response.Write"<TD vAlign=top width=13 background=images/main_left_33.gif>&nbsp;</TD></TR></TBODY></TABLE><TABLE cellSpacing=0 cellPadding=0 width=162 border=0>"
Response.Write"<TBODY><TR><TD><IMG height=8 src=images/main_left_38.gifwidth=162></TD></TR></TBODY></TABLE><br> <IMG height=19 src=images/main_left_34.gif width=162></TD>"
Response.Write"</TR></TBODY></TABLE><TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR><TD width=9 background=images/main_left_32.gif>&nbsp;</TD>"
Response.Write"<TD vAlign=top width=140 background=images/main_left_39.gif height=30> <P><IMG height=24 src=images/main_left_37.gif width=101></P></TD><TD width=13 background=images/main_left_33.gif>&nbsp;</TD></TR></TBODY></TABLE>"
Response.Write"<TABLE cellSpacing=0 cellPadding=0 width=162 border=0><TBODY><TR><TD width=9 background=images/main_left_32.gif>&nbsp;</TD><TD vAlign=top width=140 background=images/main_left_39.gif> <DIV align=center><TABLE cellSpacing=0 cellPadding=2 width=140 border=0><TBODY>"
Response.Write"<TR><TD><P><script language=javascript src=sm.js></script></P></TD></TR></TBODY></TABLE></DIV></TD><TD width=13 background=images/main_left_33.gif>&nbsp;</TD></TR></TBODY></TABLE><IMG height=8 src=images/main_left_38.gif width=162></td>"
Response.Write"<td width=79% align=left valign=top> <table width=100% border=0><tr><td colspan=2><table width=100% border=0 cellpadding=0 cellspacing=0><tr>"
Response.Write"<td background=images/newtop_6_03.gif>&nbsp;&nbsp;&nbsp;<a href=mysp.asp>我的物品</a>&nbsp;&nbsp;&nbsp;<a href=myphoto.asp>我的像册</a>&nbsp;&nbsp;&nbsp; <strike>购买记录</strike>&nbsp;&nbsp;&nbsp;<strike>赠送记录</strike>&nbsp;&nbsp;&nbsp; <strike>获赠记录</strike>&nbsp;&nbsp;&nbsp; <a href=# onClick=javascript:faceurl() >形像代码</a></td></tr><tr>"
Response.Write"<td><marquee scrollamount=6 scrolldelay=100>"& Application("jhshowgg") &"</marquee></td></tr></table></td></tr><tr>"
Response.Write"<td height=131><table width=290 border=0 align=center cellpadding=0 cellspacing=0><tr><td colspan=3><img height=39 src=images/gg01.gif width=290></td></tr><tr>"
Response.Write"<td width=50><img height=72 src=images/gg02.gif width=50></td><td valign=top width=214 bgcolor=#ffffff> <table cellspacing=0 cellpadding=0 width=210 border=0><tr>"
Response.Write"<td width=15><img height=8 src=main1.files/gg05.gif width=8></td><td>球衣/节日礼服/汽车/风景名胜</td></tr><tr>"
Response.Write"<td width=15><img height=8 src=main1.files/gg05.gif width=8></td><td>伊拉克战火弥漫</td></tr><tr><td width=15><img height=8 src=main1.files/gg05.gif width=8></td><td>制服/雨伞</td></tr><tr>"
Response.Write"<td width=15><img height=8 src=main1.files/gg05.gif width=8></td><td>持续提示：付费、测试</td></tr></table></td>"
Response.Write"<td width=26><img height=72 src=images/gg03.gif width=26></td></tr><tr><td colspan=3><img height=18 src=images/gg04.gif width=290></td></tr></table></td>"
Response.Write"<td align=left><a href=help.htm target=_blank><img height=121 alt=江湖秀向导 src=images/ad.gif width=265 border=0></a></td></tr><tr><td height=18 colspan=2><TABLE width=98% border=0 align=center cellPadding=0 cellSpacing=0 background=images/search03.gif>"
Response.Write"<form method=GET action=main.asp?CurPage=1&sid="& sid &"&sex="& sex &" name=search_form><TR><TD width=27 align=left><IMG height=34 src=images/search01.gif width=27></TD>"
Response.Write"<TD width=161>&nbsp;请输入商品名称：</TD><TD width=99><input type=hidden value="& sid &" name=sid><INPUT maxLength=20 size=10 name=search></TD><TD width=94><SELECT name=sex><OPTION value= selected>性别不限</OPTION><OPTION value=F>男</OPTION><OPTION value=M>女</OPTION></SELECT></TD><TD width=96><SELECT name=jbmoney>"
Response.Write"<OPTION value=-1 selected>价格不限</OPTION><OPTION value=0 >免费试用</OPTION><OPTION value=10>10币以下</OPTION><OPTION value=20>20币以下</OPTION><OPTION value=30>30币以下</OPTION><OPTION value=60>60币以下</OPTION><OPTION value=100>100币以下</OPTION><OPTION value=200>200币以下</OPTION></SELECT></TD>"
Response.Write"<TD align=middle width=85><A href=javascript:start_search();><IMG height=19 src=images/search04.gif width=53 border=0></A></TD><TD align=right width=27><IMG height=34 src=images/search02.gif width=27></TD></TR></FORM></TABLE></td></tr></table>"
Response.Write"<TABLE width=98% border=0 align=center cellPadding=0 cellSpacing=0 background=images/main_right_06.gif><TR><TD width=30><IMG height=23 src=images/main_right_11.gif width=20></TD><TD width=100><FONT color=blue><IMG height=1 src=images/1.gif width=1>"
Response.Write idsm &"</FONT></TD><TD><IMG height=1 src=images/1.gif width=1><a href=main.asp?CurPage=1&sid="& sid &"&sex=F>男</a> <a href=main.asp?CurPage=1&sid="& sid &"&sex=M>女</a></TD><TD width=80><IMG height=1 src=images/1.gif width=1> <A href=index.asp target=_top><FONT color=#000000>返回首页</FONT></A></TD>"
Response.Write"</TR></TABLE><table width=98% border=0 align=center cellpadding=0 cellspacing=0 bgcolor=#d1d3d4>"
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
	CurPage = 1
Else
	CurPage = CINT(Request.QueryString("CurPage"))
End If
rs.Open idsql, conn, 1,1
if rs.eof and rs.bof then
Response.Write"<div align=center>对不起，暂时没有发现任何记录！！ </div>"
else
RS.PageSize=12'设置每页记录数,可修改
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'设置最大记录数
Dim Totalcount
Totalcount =INT(RS.recordcount)
StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount
I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (I<RS.PageSize)
p=p+1
simg=replace(showgif,"[img]",rs("i"))
name=rs("b")
par1=rs("a")
'rs("a")
par2=rs("i")
par3=rs("d")
if par3="M" then
spsex="男"
elseif par3="F" then
spsex="女"
else
spsex="男女"
end if
if rs("e")=true then
par5=1
else
par5=0
end if
Response.Write"<tr><td width=297 height=138><table width=98% border=0 align=center cellpadding=0 cellspacing=2><tr><td width=84 bgcolor=#e6e7e8 height=84 rowspan=5> <img height=84 alt=点击鼠标左键试穿 src="& simg &" width=84 border=0 onClick=javascript:CHECK('"& par1 &"','"& par2 &"','"& par3 &"','"& par5 &"')>"
Response.Write"</td><td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;"& name &"&nbsp;(<font color=blue>"& spsex &"</font>)&nbsp;"
if rs("e")=true then
	Response.Write"<img src=images\vip_gif.gif>"
end if
Response.Write"</td></tr><tr><td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;价&nbsp;钱："& rs("f") &"币/个</td></tr><tr>"
Response.Write"<td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;VIP价："& rs("g") &"币/个</td></tr><tr>"
Response.Write"<td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;过&nbsp;期：<font color=red>"& rs("h") &"</font>天 购买:"& rs("l") &"次</td></tr><tr>"
Response.Write"<td width=191 height=20 bgcolor=#e6e7e8>&nbsp;<input name=image3 type=image onClick=javascript:CHECK('"& par1 &"','"& par2 &"','"& par3 &"','"& par5 &"') src=images/main_right_28.gif alt=试穿 width=31 height=20 border=0>&nbsp;"
Response.Write"<input name=image3 type=image onClick=javascript:buy("& rs("id") &",1,"& par5 &") src=images/main_right_15.gif alt=购买 width=31 height=20 border=0>&nbsp;<input name=image22 type=image  src=images/main_right_16.gif alt=赠送功能暂未开放 width=31 height=20 border=0>"
Response.Write"&nbsp;<img height=20 alt=收藏功能暂未开放 src=images/main_right_17.gif width=31 border=0></td><td width=10 align=right valign=bottom bgcolor=#e6e7e8><img height=20 src=images/000.gif width=13></td></tr></table></td><td width=293>"
RS.MoveNext
if RS.Eof then exit do
	p=p+1
simg=replace(showgif,"[img]",rs("i"))
name=rs("b")
par1=rs("a")
'par1=rs("a")
par2=rs("i")
par3=rs("d")
if par3="M" then
spsex="男"
elseif par3="F" then
spsex="女"
else
spsex="男女"
end if
if rs("e")=true then
par5=1
else
par5=0
end if
Response.Write"<table width=98% border=0 align=center cellpadding=0 cellspacing=2><tr><td width=84 bgcolor=#e6e7e8 height=84 rowspan=5> <img height=84 alt=点击鼠标左键试穿 src="& simg &" width=84 border=0 onClick=javascript:CHECK('"& par1 &"','"& par2 &"','"& par3 &"','"& par5 &"')></td>"
Response.Write"<td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp; "& name &"&nbsp;(<font color=blue>"& spsex &"</font>)&nbsp;"
if rs("e")=true then
	Response.Write"<img src=images\vip_gif.gif>"
end if
Response.Write"</td></tr><tr><td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;价&nbsp;钱："& rs("f") &"币/个</td></tr><tr><td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;VIP价："& rs("g") &"币/个</td></tr><tr>"
Response.Write"<td height=20 colspan=2 bgcolor=#e6e7e8>&nbsp;过&nbsp;期：<font color=red>"& rs("h") &"</font>天 购买:"& rs("l") &"次</td></tr><tr><td width=185 height=20 bgcolor=#e6e7e8>"
Response.Write"&nbsp;<input name=image3 type=image onClick=javascript:CHECK('"& par1 &"','"& par2 &"','"& par3 &"','"& par5 &"') src=images/main_right_28.gif alt=试穿 width=31 height=20 border=0>&nbsp;"
Response.Write"<input name=image32 type=image onClick=javascript:buy("& rs("id") &",1,"& par5 &") src=images/main_right_15.gif alt=购买 width=31 height=20 border=0>"
Response.Write"&nbsp;<input name=image22 type=image  src=images/main_right_16.gif alt=赠送功能暂未开放 width=31 height=20 border=0>"
Response.Write"&nbsp;<img height=20 alt=收藏功能暂未开放 src=images/main_right_17.gif width=31 border=0></td>"
Response.Write"<td width=13 align=right valign=bottom bgcolor=#e6e7e8><img height=20 src=images/000.gif width=13></td></tr></table></td></tr>"
I=I+2
RS.MoveNext
Loop
Response.Write"</table><table width=98% height=34 border=0 cellpadding=0 cellspacing=0><tr><td width=27 height=34 align=left valign=top><IMG height=34 src=images/search01.gif width=27></td><td width=93% align=right background=images/search03.gif>共"& Totalcount &"条记录["
if StartPageNum>1 then
	Response.Write"<a href=main.asp?CurPage="& StartPageNum-1 &"&sid="& sid &"&sex="& sex &"><<</a>"
end if
For I=StartPageNum to EndPageNum
if I<>CurPage then
	Response.Write"<a href=main.asp?CurPage="& I &"&sid="& sid &"&sex="& sex &">"& I &"</a> "
else
	Response.Write"<b><font color=#ff0000>"& I &"</font></b> "
end if
Next
if EndPageNum<RS.Pagecount then
Response.Write"<a href=main.asp?CurPage="& EndPageNum+1 &"&sid="& sid &"&sex="& sex &">>></a> "
end if
Response.Write"]</td><td width=3% align=right valign=top><IMG height=34 src=images/search02.gif width=27></td></tr></table>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write"</td></tr><tr align=center ><td height=12 colspan=2 valign=top background=images/newtop_5_03.gif>Copyright right Joybar 2005-2006 程序：永不放弃 声明：所有图片资料归原作者Tencent所有! </td></tr></table></BODY></HTML>"
Response.Write"<script language=JavaScript>function buy(id,sl,ishy)"
Response.Write"{if (ishy==1 && myhy==0){alert('对不起，您非会员，此装备为仅会员专用!');return false;}parent.myjhshow.location.href='buy.asp?id='+id+'&sl='+sl;}"
Response.Write"function setcookie(name){document.cookie = 'jhshow=' + name + ';';}"
Response.Write"function getcookie(name){var cname = name + '=';var dc = document.cookie;if (dc.length > 0) {begin = dc.indexOf(cname);if (begin != -1) {begin += cname.length;end = dc.indexOf(';', begin);if (end == -1) end = dc.length;return dc.substring(begin, end); }}return null;}"
Response.Write"function initshow(){if (par4=='M'){itemnos='||||||14|13|12||11||10|9||||8|||||||';}else{itemnos='||||||7|6|5||4||3|2||||1|||||||';}var itemnos=myshow;var showArray = itemnos.split('|');var s='';"
Response.Write"for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
Response.Write"s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';}}"
Response.Write"s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"&chr(34)&">';"
Response.Write"show.innerHTML=s;var newequip = '';for (var i=0; i<25; i++){temp = showArray[i];if (i==0){newequip=temp;}else{newequip += '|'+temp+'';}}"
Response.Write"setcookie(newequip);if(confirm('您是否也要将数据库中当前形像初始化？')){parent.myjhshow.location.href='initface.asp';}else{return s;}}"
Response.Write"function CheckIllegal(showArray){if((showArray[6] != '') && (showArray[7] != '') && (showArray[8] != '') && (showArray[10] != '') && (showArray[12] != '') && (showArray[13] != '') && (showArray[17] != ''))return 0;elsereturn -1;}"
Response.Write"function CHECK(par1,par2,par3,par5){if (par5==1 && myhy==0){alert('对不起，您非会员，此装备为仅会员专用!');return false;}if ((par3 != par4) && (par3 != 'U')){alert('对不起，性别不符，您无法试穿！');return false;}"
Response.Write"if (parent.old1==par1 && parent.old2==par2 && parent.old3==par3){parent.old1='';parent.old2='';parent.old3='';tuotz(par1);brushshow();return 0;}else{parent.old1=par1;parent.old2=par2;parent.old3=par3;}"
Response.Write"tuotz(par1);if (par1.indexOf('_')!=0){var taozhuang = par1.split('_');for(ii=0;ii<taozhuang.length;ii++){chuan(taozhuang[ii],par2);}}else{"
Response.Write"chuan(par1,par2);}}"
Response.Write"function tuotz(par1){if (par1.indexOf('_')!=0){var taozhuang = par1.split('_');for(ii=0;ii<taozhuang.length;ii++){tuotz1(taozhuang[ii]);}}else{tuotz1(par1);}}"
Response.Write"function tuotz1(par1){var smc=par1-1;var nowpar=showArray[smc];if (nowpar != '' && nowpar != null){var count=0;var tuotype='';for (var x=0; x<25; x++){if (showArray[x] == nowpar)"
Response.Write"{if(count == 0){tuotype = (x+1);}else{tuotype +='_' + (x+1);}count ++;}}if (count>1){if (tuotype.indexOf('_')!=0){var tuo = tuotype.split('_');for(xx=0;xx<tuo.length;xx++){switch(tuo[xx]){"
Response.Write"case '7':if(par4=='F'){showArray[tuo[xx]-1] ='7';break;}else{showArray[tuo[xx]-1] ='14';break;}case '8':if(par4=='F'){showArray[tuo[xx]-1] ='6';break;}else{showArray[tuo[xx]-1] ='13';break;}case '9':if(par4=='F'){showArray[tuo[xx]-1] ='5';break;}else{showArray[tuo[xx]-1] ='12';break;}case '11':if(par4=='F'){showArray[tuo[xx]-1] ='4';break;}else{showArray[tuo[xx]-1] ='11';break;}case '13':if(par4=='F'){showArray[tuo[xx]-1] ='3';break;}else{showArray[tuo[xx]-1] ='10';break;}"
Response.Write"case '14':if(par4=='F'){showArray[tuo[xx]-1] ='2';break;}else{showArray[tuo[xx]-1] ='9';break;}case '18':if(par4=='F'){showArray[tuo[xx]-1] ='1';break;}else{showArray[tuo[xx]-1] ='8';break;}default:showArray[tuo[xx]-1] ='';}}}}else{switch(par1){case '7':if(par4=='F'){showArray[smc] ='7';break;}else{showArray[smc] ='14';break;}"
Response.Write"case '8':if(par4=='F'){showArray[smc] ='6';break;}else{showArray[smc] ='13';break;}case '9':if(par4=='F'){showArray[smc] ='5';break;}else{showArray[smc] ='12';break;};case '11':if(par4=='F'){showArray[smc] ='4';break;}else{showArray[smc] ='11';break;}case '13':if(par4=='F'){showArray[smc] ='3';break;}else{showArray[smc] ='10';break;}"
Response.Write"case '14':if(par4=='F'){showArray[smc] ='2';break;}else{showArray[smc] ='9';break;}case '18':if(par4=='F'){showArray[smc] ='1';break;}else{showArray[smc] ='8';break;}default:showArray[smc] ='';}}}}"
Response.Write"function chuan(par1,par2){var smc=par1-1;showArray[smc] = par2;brushshow();}"
Response.Write"function brushshow(){var s='';var j=0;for (var i=0; i<25; i++){if(showArray[i] != '' && showArray[i] !=null){j=i+1;"
Response.Write"s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';}}"
Response.Write"var newequip = '';for (var i=0; i<25; i++){temp = showArray[i];if (i==0){newequip=temp;}else{newequip += '|'+temp+'';}}setcookie(newequip);"
Response.Write"s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &">';show.innerHTML=s;}"
Response.Write"var old1='';old2='';old3='';if (getcookie('jhshow')==null){setcookie(myshow);var itemnos = getcookie('jhshow');}else{var itemnos = getcookie('jhshow');}"
Response.Write"var showArray = itemnos.split('|');var s='';for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
Response.Write"s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';}}"
Response.Write"s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &">';show.innerHTML=s;"
Response.Write"function start_search(){document.search_form.submit();}"
Response.Write"function deleteshow(){document.cookie = 'jhshow=;';}"
Response.Write"function userAv(){parent.myjhshow.location.href='saveshow.asp?act=save&data='+getcookie('jhshow');}"
Response.Write"function faceurl(){window.open('myface.asp','myface','width=500,height=280,resizable=0,scrollbars=0,menubar=0,status=1');}"
Response.Write"function userPhoto(){parent.myjhshow.location.href='saveshow.asp?act=savephoto&data='+getcookie('jhshow');}</script>"

%>
