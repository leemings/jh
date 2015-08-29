<!--
var Mouse_Obj="none";
var pX
 document.onmousemove=D_NewMouseMove;
 document.onmouseup=D_NewMouseUp;
 function m(c_Obj)
 {
  Mouse_Obj=c_Obj;
  pX=parseInt(document.all(Mouse_Obj).style.left)-event.x;
  //pY=parseInt(document.all(Mouse_Obj).style.top)-event.y; 
  }
 
 function D_NewMouseMove()
 {
 if(Mouse_Obj!="none")
  {
  
  switch (Mouse_Obj)
  {
  case "Layer1":
  //alert("hello");
  //if (((pX+event.x)<480)&&((pX+event.x)>190))
  if (((pX+event.x)<230)&&((pX+event.x)>160))
  
  document.all(Mouse_Obj).style.left=pX+event.x;
  break;
  
  case "Layer2":
 // if (((pX+event.x)>275)&&((pX+event.x)<315))
 
  if (((pX+event.x)>145)&&((pX+event.x)<193))
  document.all(Mouse_Obj).style.left=pX+event.x;
  break;
  default:
  break;
  }
  
  event.returnValue=false;
  }
  }
 
 function D_NewMouseUp()
 {
 if(Mouse_Obj!="none")
  {
  
  
  switch (Mouse_Obj)
  {
  case "Layer1":
  
  MediaPlayer1.controls.CurrentPosition=(((Layer1.style.left.slice(0,3)-160)/70)*MediaPlayer1.currentMedia.duration);
  break;
  case "Layer2":
 MediaPlayer1.settings.Volume=((Layer2.style.left.slice(0,3)-145)/50)*100;
 
  break;
  
  default:
  
  break;
  }
  Mouse_Obj="none";

  }
  
  }
  
var img_flag=1;
var play_flag=0;

function song_select_color(id,list_num)
{
for(i=0;i<list_num;i++) list.document.all.item("list_tr"+i).className =" ";

list.document.all.item("list_tr"+id).className ="tr_on";
parent.play_bank.window.location.href ="song_detail.php?song_id="+MediaPlayer1.currentPlaylist .Item(id).getItemInfo ("COPYRIGHT")+"&base_num="+MediaPlayer1.currentPlaylist .Item(id).getItemInfo ("author");

}

function song_select(id,list_num)
{
MediaPlayer1.controls.playItem (MediaPlayer1.currentPlaylist .Item (id));

}

function window_loc(index_id,act)
{
//alert (current_index);
window.list.location.href="submit.php?index_id="+index_id+"&action="+act;
}

function player_hidden(rec)
{
if (rec==1)
{
MediaPlayer1.style.height="0";
MediaPlayer1.style.width="0";
document.all .item ("play_bank").style .visibility ="visible";
}
else
{
MediaPlayer1.style.height="255";
MediaPlayer1.style.width="350";
document.all .item ("play_bank").style .visibility ="hidden";
}
}
var the_time;

function song_result()
{
var media_cof;
var media_id;
   for (i=0;i<MediaPlayer1.currentPlaylist.Count;i++)
   {
   media_cof = MediaPlayer1.currentMedia .isIdentical (MediaPlayer1.currentPlaylist .Item (i));
   if (media_cof==true) media_id=i;
   
window.clearTimeout (the_time);
}

name_list.innerHTML=MediaPlayer1.currentPlaylist.Item (media_id).name;

if (document.all.item("list")&&list.document.all.item("list_tr0")) song_select_color(media_id,MediaPlayer1.currentPlaylist.Count);

}
function show_buffer()
{
bfprogress = "<font size=2 color=#FFFFFF>"+MediaPlayer1.network .bufferingProgress  +"%</font>";
buffer_list.innerHTML=bfprogress;
time_list.innerHTML="<strong><font size=4 color=#FFFFFF>"+MediaPlayer1.controls.currentPositionString+"</font></strong>";
//≥ı ºªØx
Layer1.style.left=(MediaPlayer1.controls.CurrentPosition/MediaPlayer1.currentMedia.duration)*70+160;
the_time=window.setTimeout('show_buffer()',1000);
}


function img_change_over(img_id)
{
img_id.src="images/player/"+img_id.id+"1.gif";
}

function img_change_out(img_id)
{
img_id.src="images/player/"+img_id.id+"2.gif";
} 

function change_img(img_id)
{
switch (img_flag)
{
	case 1:
	
	img_id.src="images/ccc.gif";
	player_hidden(img_flag);
	img_flag=0;
	break;
	case 0:
	img_id.src="images/ccc1.gif";
	player_hidden(img_flag)
	img_flag=1;
	 break;
	default:
	 break;
	
	}
}
function play_change_over(img_id)
{
if (play_flag==1) img_id.src="images/player/play1.gif";
else img_id.src="images/player/paused1.gif"; 
}
function play_change_out(img_id)
{
if (play_flag==1) img_id.src="images/player/play.gif";
else img_id.src="images/player/paused.gif"; 
}
function play_control(img_id)
{
	switch (img_id.id)
{
	case "play":
	if (play_flag==1)
	{
		play_flag=0;
		MediaPlayer1.controls.play();
		
		img_id.src="images/player/paused.gif";
	}else
	{
		play_flag=1;
		MediaPlayer1.controls.pause();
		img_id.src="images/player/play.gif";
		}
	
	break;
	case "stop":
	play_flag=1;
	MediaPlayer1.controls.stop();
	play.src="images/player/play.gif";
 	break;
	case "go_start":
	MediaPlayer1.controls.previous();
 	break;
	case "go_end":
	MediaPlayer1.controls.next();
	
	default:
	break;
	
	}

}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

-->