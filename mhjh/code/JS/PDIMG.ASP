<head>
<script language=javascript>
var imagedb=new Array(53);
var timetmp=10;
for(var i=0;i<=6;i++){
imagedb[i]=new Image(40,40);
imagedb[i].src='../image/dice/d'+i+'.gif';
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
<link rel=stylesheet href='../style.css'>
</head>
<body  topmargin=0 leftmargin=0 background='../chatroom/bg1.gif'>
<p align=center><font color=0000ff size=5>&nbsp;赌骰子&nbsp;</font></p>
1：三个骰子中与其余两个点数不同者为所得点数。例X+X+Y，则得点为Y，点多者胜，赔率为2<br>
2：若三个骰子得点相同则赔率升为10，或庄家与闲家一样三个骰子相同点多者胜，若再同点庄家胜<br>
3：若你没有掷中X+X+Y样式的骰子，你还可以再掷，直到三次机会用完为止。若掷中则本局结束！<br>
4：十赌九诈，劝君三思！<br>
<table border=0 cellspacing=0 width=80% cellpadding=0>
<tr ><td>庄家：小鱼儿</td><td><img border="0" src="../Image/DICE/D0.GIF"></td><td><img border="0" src="../Image/DICE/D0.GIF"></td><td><img border="0" src="../Image/DICE/D0.GIF"></td></tr>
<tr ><td>闲家：<%=session("yx8_mhjh_username")%></td><td><img border="0" src="../Image/DICE/D0.GIF"></td><td><img border="0" src="../Image/DICE/D0.GIF"></td><td><img border="0" src="../Image/DICE/D0.GIF"></td></tr>
</table>
</body>

