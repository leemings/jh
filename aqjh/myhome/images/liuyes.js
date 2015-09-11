/***/
var theLiuyeId = 0; 
var 	lyproperty;
var  lyliuyes;
//=================================methods
function getLiuYe(num){
	var  paper=lyliuyes[num];
	var s=paper.content+"<br><br>";
    s+="<a href=/app/data/tribe.hut.hut?"+paper.writer+" target=_blank>"+paper.writer+"("+paper.byname+")</a>&nbsp;"+formatTime(paper.postTime)+"&nbsp;留";
	return s;
}//  show the message
function showLiuYe(){
	if(lyliuyes.length==0) return ;
document.all.LiuyeMessage.innerHTML=getLiuYe(theLiuyeId);
document.all.liuyeNumId.innerHTML=(theLiuyeId+1)+"/"+lyliuyes.length;
}

// Called by >>> button (display next message) .
function liuyeNext() 
{
     if(lyliuyes.length<1) return;
	 if(theLiuyeId==lyliuyes.length-1) {
		 //alert("返回第一条");
		 theLiuyeId=0;
	}
	 else{
		 theLiuyeId++;
	 }
	 showLiuYe();
}
// Called by <<< button (get previous message).
function liuyePrevious() 
{
     if(lyliuyes.length<1) return;
	 if(theLiuyeId==0) {
		 //alert("跳到最后一条");
		 theLiuyeId=lyliuyes.length-1;
	}
	 else{
		 theLiuyeId--;
	 }
	 showLiuYe();
}
//  post the message to the  destinnation
function liuyeWrite(){
   document.all.LiuyeMessage.innerHTML="<input type=hidden  name=host ><input type=hidden  name=act value=add><TEXTAREA  name=content rows=8 wrap=yes cols=42 ></TEXTAREA>";
  document.all.LiuyePost.host.value=lyproperty.host;
}

//  delete the message
function liuyeDel(){
	if(lyproperty.host!=lyproperty.user) {
		alert("你不是主人，这样做不好"+lyproperty.host+" "+lyproperty.user);
		return;
	}
	document.all.LiuyeMessage.innerHTML="";
	window.open("/app/data/myinfo.LiuYeUpdate?act=del&id="+lyliuyes[theLiuyeId].id,"InfoRefresh");
}
//  refresh the datashow
function liuyeFresh(){
	document.all.LiuyeMessage.innerHTML="";
    window.open("/app/data/myinfo.LiuYeShow?type=htmjs&host="+lyproperty.host,"InfoRefresh");
}
//  refresh the datashow
function liuyePost(){
	var obj=document.all["LiuyePost"];
	if(obj==null||obj.elements[2].value==null||obj.elements[2].value==""||obj.elements[2].value.length>200) {
		alert("请填写300以内的字再提交");
		return;
	}
	obj.submit();//window.open("/app/data/myinfo.LiuYeShow?type=htmjs&host="+lyproperty.host,"InfoRefresh");
}
//===========主体方法
// 显示在层里的内容字符串
function liuyeDisplyer() {
	    var s="<div id=\"MyLiuYe\" style=\"position:absolute;visibility:;display:none; top:200px; left:225px; z-index:1\"  candrag><FORM name=LiuyePost method=post action=/app/data/myinfo.LiuYeUpdate target=InfoRefresh >	<div id=LiuyeMessage  style=\"position:absolute; left: 10px; top: 52px;width:325px;height:140px;FONT-SIZE: 12px; SCROLLBAR-FACE-COLOR: #E0F7EF;SCROLLBAR-TRACK-COLOR: #E0F7EF;SCROLLBAR-BASE-COLOR: #DFE7DF;SCROLLBAR-DARKSHADOW-COLOR:#E0F7EF;SCROLLBAR-3DLIGHT-COLOR: #E0F7EF;SCROLLBAR-HIGHLIGHT-COLOR: #808880;overflow:auto ;\"></div><img src=/myinfo/liuye/liuye.gif usemap=#liuyemap border=0><map name=liuyemap><area shape=\"rect\" coords=\"4,29,71,50\" onclick=\"liuyeWrite();\"><area shape=\"rect\" coords=\"10,193,37,208\" onclick=\"liuyeFresh();\"><area shape=\"rect\" coords=\"52,193,81,208\" onclick=\"liuyeDel();\"><area shape=\"rect\" coords=\"94,192,121,207\" onclick=\"liuyePost();\"><area shape=\"rect\" coords=\"137,192,163,208\" onclick=\"closeLayer('MyLiuYe');\"><area shape=\"circle\" coords=\"276,201,7\" onclick=\"liuyePrevious();\"><area shape=\"circle\" coords=\"323,200,7\" onclick=\"liuyeNext();\"> <area shape=\"rect\" coords=\"79,6,329,46\" dragarea style=\"cursor:hand;\" ></map><div id=\"liuyeNumId\" style=\"position:absolute; width:29px; height:14px; z-index:1; left: 285px; top: 195px;font-size:9pt;\"></div></form></div>";
        return s;
}
function  initLiuyeShow(data){
   var theLiuyes=new LiuYeS(data); 
   if(data.length>30){
	lyproperty=theLiuyes.property;
       lyliuyes=theLiuyes.liuyes;
      theLiuyeId=lyliuyes.length-1;
      showLiuYe();
   }
   else{
	   lyproperty=theLiuyes.property;
	   theLiuyeId=0;
       lyliuyes=new Array();
		document.all.liuyeNumId.innerHTML="0/0";
   }
}
/*if(document.all["LiuyePanel"]==null){
	var divS="<div id=LiuyePanel style=\"position:absolute; visibility:hidden; left: 0px; top:10px;  layer-background-color: #dddddd; border: 1px none #000000 \"></div>";
     document.writeln(divS);
	}*/
document.write(liuyeDisplyer());
