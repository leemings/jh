<!--
a = 5
var slump = Math.random(); 
var talet = Math.round(slump * (a-1))+1; 
function create() { 
this.under = '' 
}
b = new Array() 
for(var i=1; i<=a; i++) {
 b[i] = new create();
 if (i<10) {
 b[i].under = '<a href=plmm.asp target=_blank><img src=images/img_ad/0'+i+'.gif border=0></a>' ;
    }else {
 b[i].under = '<a href=plmm.asp target=_blank><img src=images/img_ad/'+i+'.gif border=0></a>' ;
 }
 } 
var visa = ""; 
document.write(b[talet].under); 
//--> 