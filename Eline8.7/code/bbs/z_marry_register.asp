<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_marry_conn.asp"-->
<%
stats="�Y������"
call nav()
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾����������õ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	call marry_nav()
	call main()
end if
call activeonline()
call footer()

sub main()%>

<SCRIPT>
var arPrice=new Array();
var arFieldPrice=new Array();
arPrice[0]=0;
arPrice[1]=0.6;
arPrice[2]=1;
arPrice[3]=1.4;
arPrice[4]=1.8;
arPrice[5]=2.2;
arPrice[6]=2.6;
arPrice[7]=3;
arPrice[8]=3.4;
arPrice[9]=3.8;
arFieldPrice[0]=0.3;
arFieldPrice[1]=0.3;
arFieldPrice[2]=0.3;
arFieldPrice[3]=0.3;
arFieldPrice[4]=0.3;
arFieldPrice[5]=0.3;
arFieldPrice[6]=0.3;
arFieldPrice[7]=0.3;
arFieldPrice[8]=0.3;
arFieldPrice[9]=0.3;
arFieldPrice[10]=0.6;
arFieldPrice[11]=0.6;
arFieldPrice[12]=0.6;
arFieldPrice[13]=0.6;
arFieldPrice[14]=0.6;
arFieldPrice[15]=0.6;
arFieldPrice[16]=0.6;
arFieldPrice[17]=0.6;
arFieldPrice[18]=0.6;
arFieldPrice[19]=1;
arFieldPrice[20]=1;
arFieldPrice[21]=1;
arFieldPrice[22]=1;
arFieldPrice[23]=1;

function showPrice()
{
	var StartHour=document.regform.wed_b_h.value;
	var LastHour=document.regform.sLastTime.value;
	var Standard=document.regform.iStandard.value;
	var FieldPrice;
	var MarryPrice;
	var total;

	if(parseInt(StartHour)+parseInt(LastHour)<24)
	{	
		//alert(parseInt(StartHour)+1)
		//alert(parseInt(StartHour)+parseInt(LastHour));
		
		FieldPrice=0;
		for(i=parseInt(StartHour)+1;i<=parseInt(StartHour)+parseInt(LastHour);i++)	
			FieldPrice+=arFieldPrice[i];
		//alert(FieldPrice);

	}
	else
	{
		//alert(parseInt(StartHour)+1)
		//alert(parseInt(StartHour)+parseInt(LastHour));
		
		FieldPrice=0;
		for(i=parseInt(StartHour)+1;i<=23;i++)	
			FieldPrice+=arFieldPrice[i];
		for(i=0;i<=parseInt(StartHour)+parseInt(LastHour)-24;i++)
			FieldPrice+=arFieldPrice[i];

		//alert(FieldPrice);

	}

	MarryPrice=arPrice[Standard];

	//alert(MarryPrice);

	total=MarryPrice+FieldPrice;

	
	//document.all['price'].innerHTML=Math.round(total*100)/100+" ���";
	document.all['price'].innerHTML=Math.round(total*100*0.8)+"���(8��)";
}

		
function MakeArray(n)
{
	this.length = n
	for (var i = 1; i <= n; i++) 
		this[i] = 0 
	return this
}

function isnumber(c)
{
	if ((c>='0') && (c<='9'))
		return true;
	else
		return false;
}

function ischar(c)
{
	if (((c>='a') && (c<='z')) || ((c>='A') && (c<='Z')))
		return true;
	else
		return false;
}

function checknumber(s)
{
	for (i=0; i<s.length; i++)
	{
		n = s.substr(i, 1)
		if (!(isnumber(n)))
		{
			return false;
		}
	}
	return true;
}


function validateDay(yearStr, monthStr, dayStr)
{
        var yearInt = parseInt(yearStr);
	var monthInt = parseInt(monthStr) - 1;
        var dayInt = parseInt(dayStr);
	if (monthInt > 11)
	{
		return false;
	}
        monthDays = new MakeArray(12)
        monthDays [0] = 31;
        monthDays [1] = 28;
        monthDays [2] = 31;
        monthDays [3] = 30;
        monthDays [4] = 31;
        monthDays [5] = 30;
        monthDays [6] = 31;
        monthDays [7] = 31;
        monthDays [8] = 30;
        monthDays [9] = 31;
        monthDays [10] = 30;
        monthDays [11] = 31;

        if (yearInt % 100 == 0)
        {
          if (yearInt % 400 == 0)
          {
            monthDays[1] = 29;
          }
        }
        else
        {
          if (yearInt % 4 == 0)
          {
            monthDays[1] = 29;
          }
        }

        if (dayInt > monthDays[monthInt])
        {
          return false;
        }
        return true;
}

function validate()
{
	var sBride = document.regform.sBride.value;
	if(sBride.length<=0)
	{
		alert("��Ҫ��˭��飿");
		document.regform.sBride.focus();
		return;
	}

	var sLoveWord = document.regform.sLoveWord.value;
	if(sLoveWord.length<=0 ||sLoveWord.length>=50)
	{
		alert("��д������������,50������");
		document.regform.sLoveWord.focus();
		return;
	}

	var bYear = document.regform.wed_b_y.value;

	var bMonth = document.regform.wed_b_m.value;

	var bDay = document.regform.wed_b_d.value;

	if (!validateDay(bYear, bMonth, bDay))
	{
			alert("����ȷ��д��ʼ����");
			document.regform.wed_b_y.focus();
			return;
	}

	var iStandard= document.regform.iStandard.selectedIndex;
	if(iStandard<=0)
	{
		alert("��ѡ�������");
		document.regform.iStandard.focus();
		return;
	}

	var cPayMethod= document.regform.cPayMethod.selectedIndex;
	if(cPayMethod<=0)
	{
		alert("��ѡ�񸶷ѷ�ʽ");
		document.regform.cPayMethod.focus();
		return;
	}
	
	document.regform.submit();
	
}
</SCRIPT>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
  <TBODY> 
  <TR>
    <TD bgColor=#FFFFFF height=465 vAlign=top width=227>
      <img border="0" src="images/img_marry/j01.jpg"></TD>
    <TD bgColor=#FFFFFF height=465 vAlign=top width=533 background="images/img_marry/b01.jpg">
    <img border="0" src="images/img_marry/j02.jpg"><BR>
      <TABLE align=center bgColor=#fff4e8 border=0 cellPadding=0 cellSpacing=0 
      width=413>
        <TBODY>
        <TR>
          <TD width=381>
            <TABLE align=center bgColor=#fff4e8 border=0 cellPadding=2 
            cellSpacing=2 class=ct-def2 width=417>
              <FORM action=z_marry_save.asp class=ct-def2 
              method=post name=regform>
              <TBODY>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD colSpan=4 height=68 bgcolor="#FFF7FF"><FONT 
                  color=#ff3300><%=membername%>��ϲ������̤�����ĵ��á�</FONT><FONT 
                  color=#000000><BR>��������¸��ע��Ŷ�������û��������������������û�д�Ӧ�㣬<br>
                    ��ô�ܱ�Ǹ���ǲ�����ǿ����¡�˭��Ļ���ô����˭���Ǽǣ�����ϵͳ���ܾ��������<BR>
                    <FONT 
                  color=#ff0000>����Ҫ�ģ���Ҫ���Է����ǳ�Ŷ��</FONT></FONT> 
                  <HR SIZE=1>
                </TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=2 width=93 bgcolor="#FFF7FF">
                  <DIV align=left><FONT color=#ff3300>�Է��ǳ�:</FONT></DIV></TD>
                <TD colSpan=3 height=2 width=310 bgcolor="#FFF7FF">
                  <DIV align=left><INPUT maxLength=20 name=sBride size=15> 
                </DIV></TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=2 width=93 bgcolor="#FFF7FF">
                  <DIV align=left><FONT color=#ff3300>��������:</FONT></DIV><BR>(�벻Ҫ����50���֣�</TD>
                <TD bgColor=#FFF7FF colSpan=3 height=2 width=310>
                  <DIV align=left>
                  <P><TEXTAREA cols=45 name=sLoveWord rows=3 wrap=VIRTUAL></TEXTAREA> 
                  </P></DIV></TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=2 width=93 bgcolor="#FFF7FF">
                  <DIV align=left><FONT color=#ff3300>������:</FONT></DIV></TD>
                <TD colSpan=3 height=2 width=310 bgcolor="#FFF7FF">
                  <DIV align=left>
                  <P><SELECT name=iStandard onchange=showPrice() size=1> 
                    <OPTION selected value="">��ѡ��</OPTION> <OPTION 
                    value=1>һ��һ��</OPTION> <OPTION value=2>����˫��</OPTION> <OPTION 
                    value=3>��������</OPTION> <OPTION value=4>����ɽ��</OPTION> <OPTION 
                    value=5>����ͷ�</OPTION> <OPTION value=6>�������</OPTION> <OPTION 
                    value=7>�������</OPTION> <OPTION value=8>��ͷ����</OPTION> <OPTION 
                    value=9>�쳤�ؾ�</OPTION></SELECT> ����֪���������ҳ�棩</P></DIV></TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=2 width=93 bgcolor="#FFF7FF"><FONT color=#ff3300>����ʼʱ��:</FONT></TD>
                <TD colSpan=3 height=2 width=310 bgcolor="#FFF7FF"><SELECT name=wed_b_y size=1 value=2003> 
                    <OPTION value=2001>2001</OPTION> <OPTION 
                    value=2002>2002</OPTION> <OPTION value=2003>2003</OPTION> 
                    <OPTION value=2004>2004</OPTION> <OPTION 
                    value=2005>2005</OPTION> <OPTION value=2006>2006</OPTION> 
                    <OPTION value=2007>2007</OPTION> <OPTION 
                    value=2008>2008</OPTION> <OPTION 
                  value=2009>2009</OPTION></SELECT>�� <SELECT name=wed_b_m 
                    size=1> <OPTION selected value=1>1</OPTION> <OPTION 
                    value=2>2</OPTION> <OPTION value=3>3</OPTION> <OPTION 
                    value=4>4</OPTION> <OPTION value=5>5</OPTION> <OPTION 
                    value=6>6</OPTION> <OPTION value=7>7</OPTION> <OPTION 
                    value=8>8</OPTION> <OPTION value=9>9</OPTION> <OPTION 
                    value=10>10</OPTION> <OPTION value=11>11</OPTION> <OPTION 
                    value=12>12</OPTION></SELECT>�� <SELECT name=wed_b_d size=1> 
                    <OPTION selected value=1>1</OPTION> <OPTION 
                    value=2>2</OPTION> <OPTION value=3>3</OPTION> <OPTION 
                    value=4>4</OPTION> <OPTION value=5>5</OPTION> <OPTION 
                    value=6>6</OPTION> <OPTION value=7>7</OPTION> <OPTION 
                    value=8>8</OPTION> <OPTION value=9>9</OPTION> <OPTION 
                    value=10>10</OPTION> <OPTION value=11>11</OPTION> <OPTION 
                    value=12>12</OPTION> <OPTION value=13>13</OPTION> <OPTION 
                    value=14>14</OPTION> <OPTION value=15>15</OPTION> <OPTION 
                    value=16>16</OPTION> <OPTION value=17>17</OPTION> <OPTION 
                    value=18>18</OPTION> <OPTION value=19>19</OPTION> <OPTION 
                    value=20>20</OPTION> <OPTION value=21>21</OPTION> <OPTION 
                    value=22>22</OPTION> <OPTION value=23>23</OPTION> <OPTION 
                    value=24>24</OPTION> <OPTION value=25>25</OPTION> <OPTION 
                    value=26>26</OPTION> <OPTION value=27>27</OPTION> <OPTION 
                    value=28>28</OPTION> <OPTION value=29>29</OPTION> <OPTION 
                    value=30>30</OPTION> <OPTION value=31>31</OPTION></SELECT>�� 
                  <SELECT name=wed_b_h onchange=showPrice() size=1> <OPTION 
                    selected value=0>0</OPTION> <OPTION value=1>1</OPTION> 
                    <OPTION value=2>2</OPTION> <OPTION value=3>3</OPTION> 
                    <OPTION value=4>4</OPTION> <OPTION value=5>5</OPTION> 
                    <OPTION value=6>6</OPTION> <OPTION value=7>7</OPTION> 
                    <OPTION value=8>8</OPTION> <OPTION value=9>9</OPTION> 
                    <OPTION value=10>10</OPTION> <OPTION value=11>11</OPTION> 
                    <OPTION value=12>12</OPTION> <OPTION value=13>13</OPTION> 
                    <OPTION value=14>14</OPTION> <OPTION value=15>15</OPTION> 
                    <OPTION value=16>16</OPTION> <OPTION value=17>17</OPTION> 
                    <OPTION value=18>18</OPTION> <OPTION value=19>19</OPTION> 
                    <OPTION value=20>20</OPTION> <OPTION value=21>21</OPTION> 
                    <OPTION value=22>22</OPTION> <OPTION 
                  value=23>23</OPTION></SELECT> ʱ </TD></TR>
				  <script language="javascript">
					regform.wed_b_y.value=<%=year(date())%>
					regform.wed_b_m.value=<%=month(date())%>
					regform.wed_b_d.value=<%=day(date())+1%>
					regform.wed_b_h.value=<%=hour(now())%>
				  </script>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=24 width=93 bgcolor="#FFF7FF"><FONT color=#ff3300>�������ʱ�䣺</FONT></TD>
                <TD colSpan=3 height=24 width=310 bgcolor="#FFF7FF">
                  <P><FONT color=#ff3300><SELECT name=sLastTime 
                  onchange=showPrice() size=1> <OPTION selected 
                    value=1>1</OPTION> <OPTION value=2>2</OPTION> <OPTION 
                    value=3>3</OPTION> <OPTION value=4>4</OPTION> <OPTION 
                    value=5>5</OPTION> <OPTION value=6>6</OPTION> <OPTION 
                    value=7>7</OPTION> <OPTION value=8>8</OPTION> <OPTION 
                    value=9>9</OPTION> <OPTION value=10>10</OPTION> <OPTION 
                    value=11>11</OPTION> <OPTION value=12>12</OPTION> <OPTION 
                    value=13>13</OPTION> <OPTION value=14>14</OPTION> <OPTION 
                    value=15>15</OPTION> <OPTION value=16>16</OPTION> <OPTION 
                    value=17>17</OPTION> <OPTION value=18>18</OPTION> <OPTION 
                    value=19>19</OPTION> <OPTION value=20>20</OPTION> <OPTION 
                    value=21>21</OPTION> <OPTION value=22>22</OPTION> <OPTION 
                    value=23>23</OPTION> <OPTION value=24>24</OPTION></SELECT> 
                  </FONT>Сʱ </P></TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=50 width=93 bgcolor="#FFF7FF"><FONT color=#ff3300>���񸶷ѷ�ʽ:</FONT></TD>
                <TD colSpan=3 height=50 width=310 bgcolor="#FFF7FF"><FONT color=#ff3300><SELECT 
                  name=cPayMethod size=1><OPTION selected 
                    value="">��ѡ��</OPTION> <OPTION value=1>�Լ�֧��</OPTION> <OPTION 
                    value=2>�Է�֧��</OPTION> <OPTION value=3>˫��ƽ��</OPTION></SELECT> 
                  </FONT>������ֽ�<%=mymoney%>��</TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD height=2 width=93 bgcolor="#FFF7FF">
                  <DIV align=left><FONT color=#ff3300>����Ԥ��:</FONT></DIV></TD>
                <TD colSpan=3 height=2 width=310 bgcolor="#FFF7FF">
                  <DIV align=left id=price>
                    </DIV></TD></TR>
              <TR bgColor=#fff4e8 vAlign=top>
                <TD colSpan=4 height=2 bgcolor="#FFF7FF">
                  <DIV align=center><FONT color=#ff6600> <A href="javascript:validate()"><IMG border=0 
                  height=18 src="images/img_marry/tijiao.gif" vspace=5 
                  width=61></A> </FONT></DIV></TD></TR></FORM></TBODY></TABLE></TD>
          </TR>
        </TBODY></TABLE></TD></TR></TBODY></TABLE>

<% end sub%>

