<!--
var bsYear;  
var bsDate;  
var bsWeek;  
var arrLen=8;	//���鳤��
var sValue=0;	//���������
var dayiy=0;	//����ڼ���
var miy=0;	//�·ݵ��±�
var iyear=0;	//��ݱ��
var dayim=0;	//���µڼ���
var spd=86400;	//ÿ�������

var year1999="30;29;29;30;29;29;30;29;30;30;30;29";	//354
var year2000="30;30;29;29;30;29;29;30;29;30;30;29";	//354
var year2001="30;30;29;30;29;30;29;29;30;29;30;29;30";	//384
var year2002="30;30;29;30;29;30;29;29;30;29;30;29";	//354
var year2003="30;30;29;30;30;29;30;29;29;30;29;30";	//355
var year2004="29;30;29;30;30;29;30;29;30;29;30;29;30";	//384
var year2005="29;30;29;30;29;30;30;29;30;29;30;29";	//354
var year2006="30;29;30;29;30;30;29;29;30;30;29;29;30";

var month1999="����;����;����;����;����;����;����;����;����;ʮ��;ʮһ��;ʮ����"
var month2001="����;����;����;����;������;����;����;����;����;����;ʮ��;ʮһ��;ʮ����"
var month2004="����;����;�����;����;����;����;����;����;����;����;ʮ��;ʮһ��;ʮ����"
var month2006="����;����;����;����;����;����;����;������;����;����;ʮ��;ʮһ��;ʮ����"
var Dn="��һ;����;����;����;����;����;����;����;����;��ʮ;ʮһ;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;��ʮ;إһ;إ��;إ��;إ��;إ��;إ��;إ��;إ��;إ��;��ʮ";

var Ys=new Array(arrLen);
Ys[0]=919094400;Ys[1]=949680000;Ys[2]=980265600;
Ys[3]=1013443200;Ys[4]=1044028800;Ys[5]=1074700800;
Ys[6]=1107878400;Ys[7]=1138464000;

var Yn=new Array(arrLen);   //ũ���������
Yn[0]="��î��";Yn[1]="������";Yn[2]="������";
Yn[3]="������";Yn[4]="��δ��";Yn[5]="������";
Yn[6]="������";Yn[7]="������"; 
var D=new Date();
var yy=D.getYear();
var mm=D.getMonth()+1;
var dd=D.getDate();
var ww=D.getDay();
if (ww==0) ww="<font color=red>������";
if (ww==1) ww="����һ";
if (ww==2) ww="���ڶ�";
if (ww==3) ww="������";
if (ww==4) ww="������";
if (ww==5) ww="������";
if (ww==6) ww="<font color=#FF9900>������";
ww=ww;
var ss=parseInt(D.getTime() / 1000);
if (yy<100) yy="19"+yy;

for (I=0;I<arrLen;I++)
	if (ss>=Ys[I]){
		iyear=I;
		sValue=ss-Ys[I];    //���������
		}
dayiy=parseInt(sValue/spd)+1;    //���������

var dpm=year1999;
if (iyear==1) dpm=year2000;
if (iyear==2) dpm=year2001;
if (iyear==3) dpm=year2002;
if (iyear==4) dpm=year2003;
if (iyear==5) dpm=year2004;
if (iyear==6) dpm=year2005;
if (iyear==7) dpm=year2006;
dpm=dpm.split(";");

var Mn=month1999;
if (iyear==2) Mn=month2001;
if (iyear==5) Mn=month2004;
if (iyear==7) Mn=month2006;
Mn=Mn.split(";");

var Dn="��һ;����;����;����;����;����;����;����;����;��ʮ;ʮһ;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;��ʮ;إһ;إ��;إ��;إ��;إ��;إ��;إ��;إ��;إ��;��ʮ";
Dn=Dn.split(";");

dayim=dayiy;

var total=new Array(13);
total[0]=parseInt(dpm[0]);
for (I=1;I<dpm.length-1;I++) total[I]=parseInt(dpm[I])+total[I-1];
for (I=dpm.length-1;I>0;I--)
	if (dayim>total[I-1]){
		dayim=dayim-total[I-1];
		miy=I;
		}
bsWeek=ww;
bsDate=yy+"��"+mm+"��";
bsDate2=dd;
bsYear="ũ��"+Yn[iyear];
bsYear2=Mn[miy]+Dn[dayim-1];
if (ss>=Ys[7]||ss<Ys[0]) bsYear=Yn[7];
function CAL(){
document.write("<table border='0' cellspacing='0' width='%100' cellpadding='1' align='center'");
document.write("<tr><td align='center'><font color=#000000>"+bsDate+"</font><br><b><font face='Arial' size='4' color=#000000>"+bsDate2+"</font><br><font color=#000000><span style='FONT-SIZE: 9pt'>");document.write(bsWeek+"</span>"+"<br></b><font color=#000000>");
document.write("</td></tr></table>");
}
//-->

