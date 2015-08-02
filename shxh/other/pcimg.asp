<%
  for i=0 to 1
	msg=msg&"<tr align=center>"
	for j=0 to 5
		msg=msg&"<td><img src='../images/card/back.gif'></td>"
	 next
	 msg=msg&"</tr>"
   next
 Response.Write "<head><script language=javascript>var imagedb=new Array(53);for(var i=0;i<53;i++){imagedb[i]=new Image(71,96);imagedb[i].src='../images/card/pc'+i+'.gif';}imagedb[53]=new Image(71,96);imagedb[53].src='../images/card/back.gif';function showcard(z0,z1,z2,z3,z4,z5,u0,u1,u2,u3,u4,u5){document.images[0].src=imagedb[z0].src;document.images[1].src=imagedb[z1].src;document.images[2].src=imagedb[z2].src;document.images[3].src=imagedb[z3].src;document.images[4].src=imagedb[z4].src;document.images[5].src=imagedb[z5].src;document.images[6].src=imagedb[u0].src;document.images[7].src=imagedb[u1].src;document.images[8].src=imagedb[u2].src;document.images[9].src=imagedb[u3].src;document.images[10].src=imagedb[u4].src;document.images[11].src=imagedb[u5].src;}</script></head><body oncontextmenu=self.event.returnValue=false topmargin=0 leftmargin=0 bgcolor='"&Application("Ba_jxqy_backgroundcolor")&"' background='"&Application("Ba_jxqy_backgroundimage")&"'><p align=center><font color=0000ff size=5>疯狂21点</font></p><hr size=3>1:A,J,Q,K皆为1点，其余按牌面计点,点大者胜。<br>2:不超过21点且不超过六张牌可继续要牌，也可叫开牌<br>3：和庄家点数相同者计庄家胜。<table border=0 cellspacing=0 width=100% cellpadding=0>"&msg&"</table></body>"
%>
