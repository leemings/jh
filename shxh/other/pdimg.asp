<head>
<script language=javascript>
var imagedb=new Array(53);
var timetmp=10;
for(var i=0;i<=6;i++){
	imagedb[i]=new Image(40,40);
	imagedb[i].src='../images/dice/d'+i+'.gif';
}
function showcard(z0,z1,z2,z3,z4,z5,u0,u1,u2,u3,u4,u5){
	document.images[0].src=imagedb[z0].src;
	}
function showdice(z,z0,z1,z2){
	timetmp=timetmp+10;
	if(timetmp<=100){
	var rnd0=Math.round(Math.random()*5)+1;
	var rnd1=Math.round(Math.random()*5)+1;
	var rnd2=Math.round(Math.random()*5)+1;
	document.images[z].src=imagedb[rnd1].src;
	document.images[z+1].src=imagedb[rnd1].src;
	document.images[z+2].src=imagedb[rnd2].src;
	setTimeout('showdice('+z+','+z0+','+z1+','+z2+');',timetmp);
	}
	else{
	document.images[z].src=imagedb[z0].src;
	document.images[z+1].src=imagedb[z1].src;
	document.images[z+2].src=imagedb[z2].src;
	timetmp=10;
	}
}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 leftmargin=0 bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>' >
<p align=center><font color=0000ff size=5>������</font></p><hr size=3>
1����������������������������ͬ��Ϊ���õ�������X+X+Y����õ�ΪY�������ʤ������Ϊ2<br>
2�����������ӵõ���ͬ��������Ϊ10����ׯ�����м�һ������������ͬ�����ʤ������ͬ��ׯ��ʤ<br>
3������û������X+X+Y��ʽ�����ӣ��㻹����������ֱ�����λ�������Ϊֹ���������򱾾ֽ�����<br>
4��ʮ�ľ�թ��Ȱ����˼��<br>
<table border=0 cellspacing=0 width=80% cellpadding=0>
<tr ><td>ׯ�ң���ԯ����</td><td><img src='../images/dice/d0.gif'></td><td><img src='../images/dice/d0.gif'></td><td><img src='../images/dice/d0.gif'></td></tr>
<tr ><td>�мң�<%=session("Ba_jxqy_username")%></td><td><img src='../images/dice/d0.gif'></td><td><img src='../images/dice/d0.gif'></td><td><img src='../images/dice/d0.gif'></td></tr>
</table>
</body>

