
function start_menu()
{
this.AddExtendMenu=AddExtendMenu;
this.AddItem=AddItem;
this.GetMenu=GetMenu;
this.HideAll=HideAll;
this.I_OnMouseOver=I_OnMouseOver;
this.I_OnMouseOut=I_OnMouseOut;
this.I_OnMouseUp=I_OnMouseUp;
this.P_OnMouseOver=P_OnMouseOver;
this.P_OnMouseOut=P_OnMouseOut;
    A_rbpm = new Array();
    HTMLstr  = "";
    HTMLstr += "<!-- RightButton PopMenu -->\n";
    HTMLstr += "\n";
    HTMLstr += "<!-- PopMenu Starts -->\n";
    HTMLstr += "<div id='E_rbpm' class='rm_div'>\n";
                        // rbpm = right button pop menu
    HTMLstr += "<table width='100%' border='0' cellspacing='0'>\n";
    HTMLstr += "<tr><td height='264' width='20' valign='bottom'  bgcolor='#000000'  onclick='window.event.cancelBubble=true;' onmouseup='window.event.cancelBubble=true;'>\n";
    HTMLstr += "<img  onclick='window.event.cancelBubble=true;'' onmouseup='window.event.cancelBubble=true;'  border='0' src='icon/start_menu/welcome.gif' width='20' height='254'>\n";
    HTMLstr += "</td><td height='264' width='120' style='padding: 1' valign='bottom'>\n";
    HTMLstr += "<table width='100%' border='0' cellspacing='0'>\n";
    HTMLstr += "<!-- Insert A Extend Menu or Item On Here For E_rbpm -->\n";
    HTMLstr += "</table></td></tr></table>\n";
    HTMLstr += "</div>\n";
    HTMLstr += "<!-- Insert A Extend_Menu Area on Here For E_rbpm -->";
    HTMLstr += "\n";
    HTMLstr += "<!-- PopMenu Ends -->\n";
  }
  function AddExtendMenu(id,img,wh,name,parent)
  {
    var TempStr = "";

    eval("A_"+parent+".length++");
    eval("A_"+parent+"[A_"+parent+".length-1] = id");  // 将此项注册到父菜单项的ID数组中去
    TempStr += "<div id='E_"+id+"' class='rm_div'>\n";
    TempStr += "<table width='100%' border='0' cellspacing='0'>\n";
    TempStr += "<!-- Insert A Extend Menu or Item On Here For E_"+id+" -->";
    TempStr += "</table>\n";
    TempStr += "</div>\n";
    TempStr += "<!-- Insert A Extend_Menu Area on Here For E_"+id+" -->";
    TempStr += "<!-- Insert A Extend_Menu Area on Here For E_"+parent+" -->";
    HTMLstr = HTMLstr.replace("<!-- Insert A Extend_Menu Area on Here For E_"+parent+" -->",TempStr);

    eval("A_"+id+" = new Array()");
    TempStr  = "";
    TempStr += "<!-- Extend Item : P_"+id+" -->\n";
    TempStr += "<tr id='P_"+id+"' class='out'";
    TempStr += " onmouseover='P_OnMouseOver(\""+id+"\",\""+parent+"\")'";
    TempStr += " onmouseout='P_OnMouseOut(\""+id+"\",\""+parent+"\")'";
    TempStr += " onmouseup=window.event.cancelBubble=true;";
    TempStr += " onclick=window.event.cancelBubble=true;";
    TempStr += "><td nowrap>";
    TempStr += "<img border='0' src='icon/start_menu/"+img+".gif' align='absmiddle' hspace='2' vspace='1' width='"+wh+"' height='"+wh+"'>&nbsp;"+name+"&nbsp;&nbsp</td><td style='font-family: webdings; text-align: ;'>4";
    TempStr += "</td></tr>\n";
    TempStr += "<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->";
    HTMLstr = HTMLstr.replace("<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->",TempStr);
  }
  function AddItem(id,img,wh,name,parent,location)
  {
    var TempStr = "";
    var ItemStr = "<!-- ITEM : I_"+id+" -->";
    if(id == "sperator")
    {
      TempStr += ItemStr+"\n";
      TempStr += "<tr class='out' onclick='window.event.cancelBubble=true;' onmouseup='window.event.cancelBubble=true;'><td colspan='2' height='1'><hr class='sperator'></td></tr>";
      TempStr += "<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->";
      HTMLstr = HTMLstr.replace("<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->",TempStr);
      return;
    }
    if(HTMLstr.indexOf(ItemStr) != -1)
    {
      alert("I_"+id+"already exist!");
      return;
    }
    TempStr += ItemStr+"\n";
    TempStr += "<tr id='I_"+id+"'  class='out'";
    TempStr += " onmouseover='I_OnMouseOver(\""+id+"\",\""+parent+"\")'";
    TempStr += " onmouseout='I_OnMouseOut(\""+id+"\")'";
    TempStr += " onclick='window.event.cancelBubble=true;'";
    if(location == null)
      TempStr += " onmouseup='I_OnMouseUp(\""+id+"\",\""+parent+"\",null)'";
    else
      TempStr += " onmouseup='I_OnMouseUp(\""+id+"\",\""+parent+"\",\""+location+"\")'";
    TempStr += "><td nowrap>";
    TempStr +="<img border='0' src='icon/start_menu/"+img+".gif' align='absmiddle' hspace='2' vspace='1' width='"+wh+"' height='"+wh+"'>&nbsp;"+ name+"&nbsp;";
    TempStr += "</td><td></td></tr>\n";
    TempStr += "<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->";
    HTMLstr = HTMLstr.replace("<!-- Insert A Extend Menu or Item On Here For E_"+parent+" -->",TempStr);
  }
  function GetMenu()
  {
    return HTMLstr;
  }
  function I_OnMouseOver(id,parent)
  {
    var Item;
    if(parent != "rbpm")
    {
      var ParentItem;
      ParentItem = eval("P_"+parent);
      ParentItem.className="over";
    }
    Item = eval("I_"+id);
    Item.className="over";
    HideAll(parent,1);
  }
  function I_OnMouseOut(id)
  {
    var Item;
    Item = eval("I_"+id);
    Item.className="out";
  }
  function I_OnMouseUp(id,parent,location)
  {
    var ParentMenu;
    window.event.cancelBubble=true;
    OnClick();
    ParentMenu = eval("E_"+parent);
    ParentMenu.display="none";
    if(location == null)
      eval("Do_"+id+"()");
    else
      window.open(location);
  }
  function P_OnMouseOver(id,parent)
  {
    var Item;
    var Extend;
    var Parent;
    if(parent != "rbpm")
    {
      var ParentItem;
      ParentItem = eval("P_"+parent);
      ParentItem.className="over";
    }
    HideAll(parent,1);
    Item = eval("P_"+id);
    Extend = eval("E_"+id);
    Parent = eval("E_"+parent);
    Item.className="over";
    Extend.style.display="block";
    Extend.style.posLeft=document.body.scrollLeft+Parent.offsetLeft+Parent.offsetWidth-4;
    if(Extend.style.posLeft+Extend.offsetWidth > document.body.scrollLeft+document.body.clientWidth)
        Extend.style.posLeft=Extend.style.posLeft-Parent.offsetWidth-Extend.offsetWidth+8;
    if(Extend.style.posLeft < 0) Extend.style.posLeft=document.body.scrollLeft+Parent.offsetLeft+Parent.offsetWidth;
    Extend.style.posTop=Parent.offsetTop+Item.offsetTop;
    if(Extend.style.posTop+Extend.offsetHeight > document.body.scrollTop+document.body.clientHeight)
      Extend.style.posTop=document.body.scrollTop+document.body.clientHeight-Extend.offsetHeight;
    if(Extend.style.posTop < 0) Extend.style.posTop=0;
  }
  function P_OnMouseOut(id,parent)
  {
  }
  function HideAll(id,flag)
  {
    var Area;
    var Temp;
    var i;
    if(!flag)
    {
      Temp = eval("E_"+id);
      Temp.style.display="none";
    }
    Area = eval("A_"+id);
    if(Area.length)
    {
      for(i=0; i < Area.length; i++)
      {
        HideAll(Area[i],0);
        Temp = eval("E_"+Area[i]);
        Temp.style.display="none";
        Temp = eval("P_"+Area[i]);
        Temp.className="out";
      }
    }
  }
  function s_t(){
  if (startmenu.className=="down"){
	s_img.focus();
	window.status="选定菜单中的项目，单击鼠标左键。";}
else{window.status="完成";}}
  document.onmouseup=mu;
  function mu(){OnClick();D_NewMouseUp();}
  function OnMouseUp()
  {
      var PopMenu;
      PopMenu = eval("E_rbpm");
      HideAll("rbpm",0);
      PopMenu.style.display="block";
      PopMenu.style.posLeft=0;
      PopMenu.style.posBottom=25;
      startmenu.className="down";
  }
function startclick(){
	if (startmenu.className=="up") {
	startmenu.className="down";
	OnMouseUp()}
	else{
	startmenu.className="up";
	HideAll("rbpm",0);}}
  function OnClick(){
if (((document.body.clientHeight-window.event.y)>26)|(window.event.x>55))
{
HideAll("rbpm",0);
startmenu.className="up";
}}
  // Add Your Function on following
  function Do_viewcode(){window.location="view-source:"+window.location.href;}
  function Do_windows_help(){showHelp("windows.chm::/MS-ITS:ntdef.chm::/searchhelp.htm")}
  function Do_shut() {window.close();}
  function Do_refresh() {window.location.reload();}
function Do_run()
{window.open("run.asp","","top=100,left=100,width=364,height=170");}
function Do_logoff(){
open("log.htm","","top=10000,left=10000,status=no")}
function Do_help(){window.open("file.asp?id=355");}
function setico()
{
/*Haiwa@20:45 2002-8-26*/
	divCount = document.all.ico.all.tags("DIV");
	for (i=0; i<divCount.length; i++) {
	obj = divCount(i);
	win_y=document.body.clientHeight-ico_top-ico_bottom;//图标的纵坐标范围
	n=Math.round(win_y/ico_y);//每列可排列的个数
	num=i+1;
if (n!=0){
	l=Math.round(num/n+0.49);//图标的列数
	numy=i-n*l+n;//在每一列中图标的序数
	obj.style.left=l*ico_x-ico_x+ico_left;
	obj.style.top=numy*ico_y+ico_top;
	}
	}
}