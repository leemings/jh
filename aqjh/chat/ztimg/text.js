function showtip(text)  //显示链接的说明
{
  if (document.all&&document.readyState=="complete")//针对IE
  {
  //显示跑马灯,内容就是提示框的内容
  document.all.tooltip.innerHTML="<div vAlign=center>"+text+"</div>";
  document.all.tooltip.style.pixelLeft=event.clientX+document.body.scrollLeft;
  document.all.tooltip.style.pixelTop=event.clientY+document.body.scrollTop+15;
   if (event.clientX>600)
  {
  document.all.tooltip.style.pixelLeft=(screen.width-700)/2+230;
  document.all.tooltip.style.pixelTop=event.clientY+document.body.scrollTop+15;}
   document.all.tooltip.style.visibility="visible";
  }
}

document.write(
"<div id=\"tooltip\" style=\"position:absolute;visibility:hidden; padding:3px;border:1px solid #000000;\
;background-color:#004382; height: 19px; left:77;top: 96px;z-index:10;\"></div>");
function hidetip()  //隐藏链接的说明
{
if (document.all)
  document.all.tooltip.style.visibility="hidden";
}

function setcolor(fntcolor,bgcolor)
{
  document.all.tooltip.style.bgcolor=bgcolor;
  document.all.tooltip.innerHTML.Color = fntcolor;
}

function showhint(text)
{
      onmousemove=showtip(text);
      onmouseover=showtip(text);
      onmouseout=hidetip();
}
