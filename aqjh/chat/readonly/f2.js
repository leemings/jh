if(window==window.top){top.location.href="jhchat.asp";}
var max=9;var whamsg=new Array(max+1);var base=0;var p=0;var j;for(j=0;j<=max+1;j++){whamsg[j]='';}
function addOne(what){if (base<max+1){whamsg[base]=what;base++;}else{for (i=0;i<max;i++)whamsg[i]=whamsg[i+1];whamsg[i]=what;}p=base;}
function gop(){if(p>0)p--;document.af.sytemp.value=whamsg[p];document.af.sytemp.focus();}
function gon(){if(p<base)p++;document.af.sytemp.value=whamsg[p];document.af.sytemp.focus();}
function bs(){document.af.sytemp.focus();}
function bs2(){document.af.sytemp.focus();}
function dj(){document.af.towho.value="���";document.af.towho.text="���";document.af.sytemp.focus();}
function rc(list){if(list!="0"){document.af.sytemp.value=list;document.af.sytemp.focus();}
if(list=="/վ������$"){document.af.sytemp.value='';document.af.sytemp.focus();document.af.mdsx.checked=false;parent.f3.location="fafang/admin.asp";}
if(list=="/���ŷ���$"){document.af.sytemp.value='';document.af.sytemp.focus();document.af.mdsx.checked=false;parent.f3.location="fafang/admin1.asp";}
if(list=="/�ٸ��Ź�$"){document.af.sytemp.value='';document.af.sytemp.focus();document.af.mdsx.checked=false;parent.f3.location="fafang/admin2.asp";}
if(list=="/��������"){document.af.sytemp.value='';document.af.sytemp.focus();document.af.mdsx.checked=false;parent.f3.location="horse.asp";}
if(list=="/�������$"){window.open("../jhmp/mp3.asp?you="+document.af.towho.value);document.af.sytemp.value='';document.af.sytemp.focus();}
if(list=="/��������$"){window.open("../jhmp/index.asp");document.af.sytemp.value='';document.af.sytemp.focus();}
if(list=="/ָ������$"){window.open("sendmp.asp?id=1");document.af.sytemp.value='';document.af.sytemp.focus();}
if(list=="/���ŷ�Ǯ$"){window.open("sendmp.asp?id=0");document.af.sytemp.value='';document.af.sytemp.focus();}
if(list=="/�鿴��Ʒ$"){window.open("towupin.asp?toname="+document.af.towho.value);document.af.sytemp.value='';document.af.sytemp.focus();}
if(list=="/��Ϣ$"){window.open('../yamen/mt.asp?action='+document.af.towho.value,'mt','scrollbars=yes,toolbar=no,menubar=no,width=460,height=430');document.af.sytemp.value='';document.af.sytemp.focus();}
}
function runnosay(){
document.af.clock.value=document.af.clock.value-1;
if(document.af.clock.value==300){
if(document.af.addvalues.checked)		{
parent.d.location.href='autosay.asp';
startnosay();
}else{window.open('comeon.asp','','Status=no,scrollbars=no,resizable=no,width=210,height=130');}
}
if(document.af.clock.value==0){top.location.href='nosaytimeout.asp';}if(document.af.clock.value<0){startnosay();parent.rn();}setTimeout('runnosay()',1000);}
startnosay();
runnosay();