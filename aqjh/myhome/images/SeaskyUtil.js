///======seekwhere use
var places;

function  Place(s){
    var _s=s.split(":");
	this.id=_s[0];
	this.cname=_s[1];
}
function initPlaces(){
	var placedata="0:���֪��|10:���ڼ���|11:��԰�ֻ�|12:�����ż�|13:��ɽ��ˮ|14:�ɿ�����|15:ɭ�ִ���|16:ˮ�ϲ���|17:�󺣲���|18:��ͳ�|20:Ϳдʯǽ|21:��ѧԺ|30:¶������|40:���вɹ�|41:��Ʒ������|42:���ݽ���|50:���´���|51:�����|52:�齱�|60:�ݷ�Ӣ�۴�|61:��������|70:���ֲ�|80:������|1000:������";
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
	return "���֪��";
}

initPlaces();
/////////����С������
function getTownType(id){
	var  typeName=["δ����","������","������","��ɽ��","ƽԭ��","ɭ����","ˮ����"];
	return typeName[id];
}
function getWorkType(id){
	var  typeName=["��","�ɱ�","������","�ɿ�","ƽԭ��","ɭ����","ˮ����"];
	return typeName[id];
}
function getTownClass(id){
	var jbName=["","��","��","��","��","��"];
	return jbName[id];
}
var  jbtitle= ["����է��","��������","С������","����һ��","��������","һ����ʦ","������ʥ","�������","��������","���˺�һ","�������"];


/// about show playerImage 
function showPlayerImage(player,b){
	var s="";
	if(b)s+="<a href=/app/data/land.ImageUpload target=uploadplayerimage>";
	else s+="<a href=/app/data/tribe.hut.hut?"+player.name+" target=_blank>";
	s+="<img src=\"/app/data/land.PlayerImage?username="+player.name+"\" style=\"width:64px;height:64px;";
	if(player.online==0) s+="filter:gray;";
	s+="\" border=0  title=\""+player.byname;
	if(player.sex==1) s+="����";
	if(player.sex==2)s+="Ůʿ";
	if(player.member==1) s+=" ��Ա";
	if(player.online==0) s+=" ������";
	if(player.online==1) s+=" ����";
//	if(player.honor!=null&&player.honor!="") s+=" ����:"+player.honor;
//	if(player.medal!=null&&player.medal!="") s+=" ����:"+player.medal;
	if(player.bbsjb>0) {
		s+=" �Ļ�����:"+player.bbsjb;
		if(player.meili>0) s+=" ����ֵ:"+player.meili;
	}
//	if(player.town!=null&&player.town!="") {
//		s+=" "+player.town;
//		if(player.hutnumber!="") s+=player.hutnumber+"#";
//		if(player.career!="") s+=" "+player.career;
//	}
		s+="\"></a>";
	return s;
}