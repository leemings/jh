function MM_findObj(n, d) { 
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i>d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function insertTag(MyString)
{
fontbegin='[color=' + MyString + ']'
fontend='[/color]'
fontchuli();
}

var base_hexa = "0123456789ABCDEF";
function dec2Hexa(number)
{
   return base_hexa.charAt(Math.floor(number / 16)) + base_hexa.charAt(number % 16);
}

function RGB2Hexa(TR,TG,TB)
{
  return "#" + dec2Hexa(TR) + dec2Hexa(TG) + dec2Hexa(TB);
}
function lightCase(MyObject,objName)
{
	MM_findObj(objName).bgColor = MyObject.bgColor;
}

function lightCase1(MyObject,objName)
{
	MM_findObj(objName).value = MyObject.bgColor;
}

col = new Array;
col[0] = new Array(255,0,255,-1,255,-1);
col[1] = new Array(255,0,0,1,0,0);
col[2] = new Array(255,-1,255,0,0,0);
col[3] = new Array(0,0,255,0,0,1);
col[4] = new Array(0,0,255,-1,255,0);
col[5] = new Array(0,1,0,0,255,0);
col[6] = new Array(255,-1,0,0,255,-1);

function rgb(pas,w,h){
	for (j=0;j<6+1;j++){
		for (i=0;i<pas+1;i++){
			r = Math.floor(col[j][0]+col[j][1]*i*(255)/pas);
			g = Math.floor(col[j][2]+col[j][3]*i*(255)/pas);
			b = Math.floor(col[j][4]+col[j][5]*i*(255)/pas);
		  codehex = r + '' + g + '' + b;
		  document.write('<td bgColor=\"' + RGB2Hexa(r,g,b) + '\" onClick=\"insertTag(this.bgColor);lightCase(this,\'ColorUsed\');\" width=\"'+w+'\" height=\"'+h+'\"><IMG height='+h+' width='+w+' border=0 title=\"<FONT COLOR='+RGB2Hexa(r,g,b)+'>字体颜色：'+RGB2Hexa(r,g,b)+'</FONT>\"></TD>\n');
		}
	}
}

function rgb1(pas,w,h){
for (j=0;j<6+1;j++){
    for (i=0;i<pas+1;i++)
	{
	r = Math.floor(col[j][0]+col[j][1]*i*(255)/pas);
	g = Math.floor(col[j][2]+col[j][3]*i*(255)/pas);
	b = Math.floor(col[j][4]+col[j][5]*i*(255)/pas);
  codehex = r + '' + g + '' + b;
  document.write('<td bgColor=\"' + RGB2Hexa(r,g,b) + '\" onClick=\"lightCase(this,\'ColorHead\');lightCase1(this,\'ChenTopicColor\');\" width=\"'+w+'\" height=\"'+h+'\"><IMG height='+h+' width='+w+' border=0 title=\"<FONT COLOR='+RGB2Hexa(r,g,b)+'>标题颜色：'+RGB2Hexa(r,g,b)+'</FONT>\"></TD>\n');
	}}}

function search(text,caract)
{
	for(i=0;i<text.length;i++)
	{
		if (caract == text.substring(i,i+1))
		return i+1;
	}
}
