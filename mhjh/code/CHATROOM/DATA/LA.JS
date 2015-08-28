var n = (document.layers) ? 1:0; 
var ie = (document.all) ? 1:0; 
var range = ""; 
var styleObj = ""; 
var x = 0; 
var y = 0; 
var count = 0; 
var total = 6;

var count = new Array(total);
for (i=0; i<=total-1; i++) { 
  count[i] = 0; 
}

var ycenter = new Array(total);
for (i=0; i<=total-1; i++) { 
  ycenter[i] = 190; 
}

var xcenter = new Array(total);
for (i=0; i<=total-1; i++) { 
  xcenter[i] = 310; 
}

var ytheta = new Array(total);
for (i=0; i<=total-1; i++) { 
  ytheta[i] = Math.PI; 
}

var xtheta = new Array(total);
for (i=0; i<=total-1; i++) { 
  xtheta[i] = Math.PI; 
}

var yspin = new Array(total);
for (i=0; i<=total-1; i++) { 
  yspin[i] = 3.05; 
}

var xspin = new Array(total);
for (i=0; i<=total-1; i++) { 
  xspin[i] = 3.05; 
}

var yradius = new Array(total);
for (i=0; i<=total-1; i++) { 
  yradius[i] = 50; 
}

var xradius = new Array(total);
for (i=0; i<=total-1; i++) { 
  xradius[i] = 50; 
}

var xpander = new Array(total);
for (i=0; i<=total-1; i++) { 
  xpander[i] = 1; 
}

var letter = new Array(total);
for (i=0; i<=total-1; i++) { 
  letter[i] = "b" + i; 
}

if (ie) { range = "all."; 
styleObj = ".style"; } 

function shiftTo(obj, x, y) { 
  if (n) { 
    obj.moveTo(x,y); 
  } 
  else { 
    obj.pixelLeft = x; 
    obj.pixelTop = y; 
  } 
} 
  
function getObject(obj) { 
var theObj = eval("document." + range + obj + styleObj); 
return theObj; 
}  

function parabola(num) { 
var theObj = getObject(letter[num]); 
if (count[num] < 999999) { 
  x = (xradius[num]*Math.cos(xtheta[num])) + xcenter[num]; 
  y = (yradius[num]*Math.sin(ytheta[num])) + ycenter[num]; 
  shiftTo(theObj, x, y); 
  if (xpander[num]) { 
    if (count[num] % 7 == 0) { 
      yradius[num]++; xradius[num] = xradius[num]+2; 
    } 
    if (yradius[num] >= 180) xpander[num] = 0; 
  } 
  else { 
    if (count[num] % 7 == 0) { 
      yradius[num]--; xradius[num] = xradius[num]-2; 
    } 
    if (yradius[num] <= 1) xpander[num] = 1; 
  } 
  xspin[num] = xspin[num] +.001; 
  yspin[num] = yspin[num] +.0025;
  if (xspin[num] > 3.24) xspin[num]=3.04;
  if (yspin[num] > 3.24) yspin[num]=3.04;
  
  xtheta[num] = xtheta[num] + (Math.sin(xspin[num])); 
  ytheta[num] = ytheta[num] + (Math.sin(yspin[num])); 
  count[num]++; 
  setTimeout("parabola(" + num + ")",10); 
 }
} 

function init() {
  for (i=0; i<=total-1; i++) {
    setTimeout("parabola(" + i + ")",1113*i); 
  }
}