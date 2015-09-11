<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
dy=session("dy")
if IsArray(dy) then
else
Response.Write("你没开始钓鱼啊")
	Response.Redirect("../index.asp?info=你没开始钓鱼啊")
Response.end
end if
start=Request.QueryString("start")
if start=1 then
	Randomize   '初始化随机数生成器。
   	MyValue = Int((60 * Rnd) + 1)+60 
end if 
 %>
<html>
<head>
<title>甩竿及拉杆</title>
<META HTTP-EQUIV="Cache-Control:" CONTENT="no-cache, must-revalidate">                  
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript" src="layer_obj.js" type="text/JavaScript"></script>
<link href="../the9.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
	var da1,jji,jj,tq,tqi,dd,ddi;
	 da1 = new Date();
	 jji=parseInt(da1.getMonth() /3) ;
	 tqi=<%= dy(6) %>;	
 	 dd='<%= dy(2) %>';
	 bait_name='<%= dy(4) %>';
	 bait_image='../images/<%= dy(5) %>';
	 //alert (dd+bait_name+bait_image)
	 
	 if (jji==0) jj="春天";
	 if (jji==1) jj="夏天";
	 if (jji==2) jj="秋天";
	 if (jji==3) jj="冬天";
	 if (tqi==0) tq="晴天";
	 if (tqi==1) tq="多云";
	 if (tqi==2) tq="阴天";
	 
	 if (dd=="池塘") ddi=0;
	 if (dd=="小河") ddi=1;
	 if (dd=="湖泊") ddi=2;
	 if (dd=="江河") ddi=3;
	 if (dd=="近海") ddi=4;
	 if (dd=="远洋") ddi=5;
	 
	var waittime = 0;
	var lala = 0;
	var alertStr = "";
	ns4=(document.layers)?true:false;
	ie4=(document.all)?true:false;
	var time=0<%= MyValue %>;
	time+=Math.random()*30;
	var systime = 0;
   	function myint(f)
	{
		if(f<1)return '00';
		var x=parseInt(f);
		if(x<10)x='0'+x;
		return x;
	}
	function show_fish()
	{
	 showLayer('cont4');
	 setTimeout("window.location=\"index2.asp\"",2300);
	}		
    function fmttime(systime)
	{
		var hour=myint(systime/3600); 
		var minute=myint((systime%3600)/60);
		var second=myint(systime%60);
		return hour+":"+minute+":"+second;
	}		
	function mytimer()
	{
		var tmp="";
		systime++;
		waittime++;
	
	    tmp=fmttime(systime); 
		if(waittime >= time)
		{
			hideLayer('cont16');
			hideLayer('cont14');
			alertStr='<span class=ct-def3>'+tmp+"　<a href=index2.asp?la=1>拉竿</a>"+'</span>';
			hideLayer('cont6');
			showLayer('cont5');
			showLayer('cont11');
			showLayer('cont12');
			showLayer('cont13');
			showLayer('cont15');
			lala ++ ;
			if(lala == 120) 
			{
			  hideLayer('cont5');
			  hideLayer('cont15');
			  hideLayer('cont11');
			  hideLayer('cont12');
			  hideLayer('cont13');
			  showLayer('cont17');
			  showLayer('cont14');
			  showLayer('cont6');
			  alert("鱼儿溜了！！！前功尽弃！！！");
		      window.location = "index.asp?la=2";
			}
		}
		else 
		{
		    showLayer('cont16');
			alertStr = '<span class=ct-def3>'+tmp+'</span>';
		}
		
		if(ns4)
		{
			var s=document.layers['time'].document;
			s.open();
			s.write('<span class=ct-def3>'+alertStr+'</span>');
			s.close();
		}
		if(ie4)document.all['time'].innerHTML=alertStr;
		setTimeout('mytimer()',1000);		
	}	
   </script>
</head>

<body bgcolor="ffffff" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onload=mytimer();hideLayer('cont3');hideLayer('cont4');showLayer('cont6');hideLayer('cont18');showLayer('cont14');>
<table width="760" border=0 cellpadding="0" cellspacing="0" bgcolor="#E8E8F0">
  <tr>
    <td width="167" valign="top"><br>
        <table width="167" border="0" cellspacing="0" cellpadding="0" height="80">
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/cloud02.gif" width="60" height="24"></td>
            <td><br>
                <br>
                <img src="../images/cloud03.gif" width="34" height="14"></td>
          </tr>
      </table></td>
    <td> <img src="../images/ship01.gif" width="129" height="46"><img src="../images/ship02.gif" width="137" height="95" border="0"></td>
    <td valign=top><img src="../images/ship03.gif" width="50" height="60"></td>
    <td width="100%">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="90">
        <tr>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/cloud02.gif" width="60" height="24"></td>
          <td valign="bottom"><img src="../images/cloud03.gif" width="33" height="14"><br>
              <br></td>
          <td><img src="../images/cloud01.gif" width="69" height="28"><br />
              <br></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="4" valign="top" bgcolor="#D6D7EF"><img src="../images/4_4.gif" width="30" height="20"></td>
  </tr>
</table>
<script language="JavaScript">
var content1='<img src=../images/weather'+tqi+'.gif border=0>';
var content2='<img src=../images/place'+ddi+'.jpg border=0>';
var content3='<a href=javascript:show_fish()><img  src=../images/begin.gif border=0></a>';//甩竿(静)
var content4='<img  src=../images/animation.gif border=0>';//甩竿(动)
var content5='<img  src=../images/okkk.gif border=0>';//拉竿(动)
var content6='<img  src=../images/ready.gif border=0>';//垂钓(静)
var content7='<img src="../images/icoweather'+tqi+'.gif" border=0 align=absmiddle> <span class=ct-def3>江湖天气:</span>'
             +'<span class=ct-def3>'+tq+'</span>';
var content8='<img src="../images/icoseason'+jji+'.gif" border=0 align=absmiddle><span class=ct-def3>现在的季节:</span>'
             +'<span class=ct-def3>'+jj+'</span>';
var content9='<img src="../images/icoplace'+ddi+'.gif" border=0 align=absmiddle><span class=ct-def3>垂钓地点:</span>'
             +'<span class=ct-def3>'+dd+'</span>';
var content10='&nbsp;'+'<span class=ct-def3>您带的鱼饵:</span>'
             +'<span class=ct-def3>'+bait_name+'</span><img src="'+bait_image+'" border=0 align=absmiddle>';
var content11='<img  src=../images/w_zigzag001.gif border=0>';//水花1
var content12='<img  src=../images/w_zigzag002.gif border=0>';//水花2
var content13='<img  src=../images/w_zigzag003.gif border=0>';//水花3
var content14='<img  src=../images/w_zigzag004.gif border=0>';//水花4
var content15='<span class="ct-def2">注意注意！<span class="newstopic-r">浮子动了</span>有东西咬钩喽，快抓住时机！！</span>';
var content16='<span class="ct-def2">注意水面上的<span class="newstopic-r">小水花</span>，鱼儿随时可能咬线哦！！</span>';
var content17='<span class="ct-def2">哎呀,鱼儿已经<span class="newstopic-r">溜</span>了！！</span>';
var content18='<span class="ct-def2">千万别傻愣着！<span class="newstopic-r">点击鱼竿</span>或<span class="newstopic-r">"甩竿"</span>，就可以等鱼儿上钩了！！</span>';
createLayer('cont2',160,210,271,110,true,content2);//地点场景
createLayer('cont1',280,160,100,72,true,content1);//天气场景
createLayer('cont3',116,320,200,230,true,content3);//甩竿前(静)
createLayer('cont4',132,320,200,210,false,content4);//甩竿(动)
createLayer('cont5',116,355,220,150,false,content5);//拉竿(动)
createLayer('cont6',116,355,220,150,false,content6);//垂钓(静)
createLayer('cont7',536,170,220,90,true,content7);//天气小图
createLayer('cont8',562,290,230,90,true,content8);//季节小图
createLayer('cont9',510,410,240,90,true,content9);//地点小图
createLayer('cont10',498,493,240,90,true,content10);//鱼饵小图
createLayer('cont11',170,360,55,26,false,content11);//水花小图1
createLayer('cont12',331,418,70,30,false,content12);//水花小图2
createLayer('cont13',351,468,63,24,false,content13);//水花小图3
createLayer('cont14',240,455,63,24,false,content14);//水花小图4
createLayer('cont15',140,555,535,26,false,content15);//提示1
createLayer('cont16',140,555,535,26,false,content16);//提示2
createLayer('cont17',140,555,535,26,false,content17);//提示3
createLayer('cont18',140,555,535,26,true,content18);//提示4
</script>
<table width="760" height="515"border="0" cellspacing="0" cellpadding="0" bgcolor="#d3d7ea" background="../images/bg.gif">
  <tr><td align="center">
  		   <table border="0" cellspacing="0" cellpadding="0">
<tr><td>
<script language=javascript>if(ns4)document.writeln("&nbsp;")</script><layer name='time' left=570 top=500></layer>
<div id="time" style='position: absolute;z-index:200;x-index:200;top:400;left:330;width:250;height:100;'><script language="JavaScript"></script></div>
</td></tr>
         </table>
  </td></tr>
</table>
</body>
</html>

</body>
</html>