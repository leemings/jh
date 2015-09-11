//=================================common method,about layer
function resetDiv(layername,x,y){
	var layerObj = document.all(layername);
	layerObj.style.pixelLeft=event.clientX+x;//-detailLayer.style.pixelWidth/2;
	layerObj.style.pixelTop=event.clientY+y;
}
function shLayers(layername){
	var layerObj = document.all(layername);
	if(layerObj.style.visibility=="") layerObj.style.visibility="hidden";
	else layerObj.style.visibility="";
}
function shLayer(n){
	if(document.all[n].style.display=="none"){
        document.all[n].style.visibility="";
		document.all[n].style.display="";
		document.all[n].style.zIndex=1;
	}
	else{
		document.all[n].style.visibility="hidden";
		document.all[n].style.display="none";
		document.all[n].style.zIndex=-1;
	}
}
function displayLayers(layername){
	var layerObj = document.all(layername);
	if(layerObj.style.display=="") {
		layerObj.style.display="none";
	}
	else {
		layerObj.style.display="";
		layerObj.style.visibility="";
	}
}
function frontLayer(n){
    var layerObj = document.all(n);
	layerObj.style.zIndex=getMaxZindex()+1;
}
function rearLayer(n){
    var layerObj = document.all(n);
	layerObj.style.zIndex=-1;
}
function openLayer(layername){
	var layerObj = document.all(layername);
		layerObj.style.display="";
		layerObj.style.visibility="";
		layerObj.style.zIndex=getMaxZindex()+1;
}
function closeLayer(n){
    var layerObj = document.all(n);
    layerObj.style.zIndex=-1;
	layerObj.style.visibility="hidden";
	layerObj.style.display="none";
}
///=================layer zIndex
function getMaxZindex(){
	var maxzindex=0;
	document.divs=document.all.tags("div");
	for(var i=0;i<document.divs.length;i++)
	maxzindex=(maxzindex<document.divs[i].style.zIndex)?document.divs[i].style.zIndex:maxzindex;
	return maxzindex;
}
function getiZindex(){
	var izindex=0;
	for(var i=0;i<document.frames.length;i++)
		izindex=(izindex<document.frames.style.zIndex)?document.frames.style.zIndex:izindex;
	return izindex;
}

//====表示层建立
function  setLayer(name,width,height,position,left,top){
    var s="<div id="+name+" style=\"position:"+position+";visibility:hidden;overflow:; width:"+width+"px; height:"+height+"px;  z-index:0; left: "+left+"px; top:"+top+"px; layer-background-color: #FFFFFF; border: 1px none #000000 \"></div>";
    document.writeln(s);
}
//  load data from a hidden frame
function loadData(name,src){
    var s="<div style=\"position:absolute; display:none;visibility:hidden;\"><iframe name="+name+" src="+src+"></iframe></div>";
    document.writeln(s);
}
function refresh(name,src){
    window.open(src,name);
}
function write2Layer(layername,content){
	var layer= document.all(layername);
    layer.innerHTML=content;
}




////////////////　移动和选择


function doMouseMove(){
	if((eldragged!=null)&&(1==event.button)){
		var mousex=event.clientX+document.body.scrollLeft;
		var mousey=event.clientY+document.body.scrollTop;
		eldragged.style.pixelLeft=mousex-eldragged.offsetx; 
		eldragged.style.pixelTop=mousey-eldragged.offsety;
		event.returnValue=false;
	}
}
function checkAtt(el,att){
	while(el!=null){
		if(el.getAttribute(att)!=null)
			return el;
		el=el.parentElement;
	}
	return null;
}
function doMouseDown(){ 
	eldragged=checkAtt(event.srcElement,"dragarea");
	eldragged=checkAtt(eldragged,"candrag");
	if(null!=eldragged){
		if(eldragged.getAttribute("toztop")!=null)
			eldragged.style.zIndex=getMaxZindex()+1;
		var mousex=event.clientX+document.body.scrollLeft;
		var mousey=event.clientY+document.body.scrollTop;
		eldragged.offsetx=mousex-eldragged.style.pixelLeft;
		eldragged.offsety=mousey-eldragged.style.pixelTop;
	}
}
function doSelect(){
	return((eldragged==null)&&checkAtt(event.srcElement,"nosel")==null);
}
top.eldragged=null;
top.document.onmousedown=doMouseDown;
top.document.onmousemove=doMouseMove;
top.document.onmouseup=new Function("eldragged=null;");
top.document.ondragstart=doSelect;
top.document.onselectstart=doSelect;
/*function  rightClick(){
	if (event.button==2&&checkAtt(event.srcElement,"canRight")!=null){
		alert(event.srcElement.id);
	}
}*/