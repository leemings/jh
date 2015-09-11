///======seekwhere use
var places;

function  Place(s){
    var _s=s.split(":");
	this.id=_s[0];
	this.cname=_s[1];
}
function initPlaces(){
	var placedata="0:云深不知处|10:呆在家里|11:花园种花|12:整理信笺|13:冰山化水|14:采矿炼铁|15:森林打猎|16:水畔采珠|17:大海捕鱼|18:快餐车|20:涂写石墙|21:文学院|30:露天聊聊|40:集市采购|41:商品交易所|42:房屋交易|50:议事大厅|51:公告板|52:抽奖活动|60:逸飞英雄传|61:足球天下|70:俱乐部|80:海天大地|1000:坐牢中";
     var _s=placedata.split("|");
	 places=new Array(_s.length);
	 for(var i=0;i<_s.length;i++){
	     places[i]=new Place(_s[i]);
      }
}
function  getWhere(id){
     for(var i=0;i<places.length;i++)
	{
		 if(places[i].id==id) return places[i].cname;
	}
	return "云深不知处";
}

initPlaces();
/////////关于小区划分
function getTownType(id){
	var  typeName=["未定居","极地区","海滨区","高山区","平原区","森林区","水畔区"];
	return typeName[id];
}
function getWorkType(id){
	var  typeName=["无","采冰","海滨区","采矿","平原区","森林区","水畔区"];
	return typeName[id];
}
function getTownClass(id){
	var jbName=["","村","镇","城","市","府"];
	return jbName[id];
}
var  jbtitle= ["初来乍到","略有所成","小有名气","名动一方","天下闻名","一代宗师","超凡入圣","初窥天道","无相无形","天人合一","破碎虚空"];


/// about show playerImage 
function showPlayerImage(player,b){
	var s="";
	if(b)s+="<a href=/app/data/land.ImageUpload target=uploadplayerimage>";
	else s+="<a href=/app/data/tribe.hut.hut?"+player.name+" target=_blank>";
	s+="<img src=\"/app/data/land.PlayerImage?username="+player.name+"\" style=\"width:64px;height:64px;";
	if(player.online==0) s+="filter:gray;";
	s+="\" border=0  title=\""+player.byname;
	if(player.sex==1) s+="先生";
	if(player.sex==2)s+="女士";
	if(player.member==1) s+=" 会员";
	if(player.online==0) s+=" 不在线";
	if(player.online==1) s+=" 在线";
//	if(player.honor!=null&&player.honor!="") s+=" 荣誉:"+player.honor;
//	if(player.medal!=null&&player.medal!="") s+=" 奖章:"+player.medal;
	if(player.bbsjb>0) {
		s+=" 文化级别:"+player.bbsjb;
		if(player.meili>0) s+=" 魅力值:"+player.meili;
	}
//	if(player.town!=null&&player.town!="") {
//		s+=" "+player.town;
//		if(player.hutnumber!="") s+=player.hutnumber+"#";
//		if(player.career!="") s+=" "+player.career;
//	}
		s+="\"></a>";
	return s;
}