<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"diaoyu/diao.asp")=1 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,״̬,����,����,����ʱ�� from �û� where ����='"&sjjh_name&"'",conn
ca=rs("״̬")
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<8 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����"&ss&"������������˽����ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������1�����Բ��ܳ������棡');window.close();}</script>"
	response.end
end if
if rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������������ǧ���Բ��ܳ������棡');window.close();}</script>"
	response.end
end if
if rs("����")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������������Բ��ܳ������棡');window.close();}</script>"
	response.end
end if

if ca="������" or ca="÷��" or ca="�����"then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Բ��Ͳ�Ҫ�ڳ����˰ɣ�');window.close();}</script>"
	response.end
end if
rs.close
conn.execute "update �û� set ����ʱ��=now()  where ����='"&sjjh_name&"'"
randomize timer
r=int(rnd*15)+1
select case r
case 1
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶������������ǿ�ϲ�ɺ�<IMG src='../work/catch/fish.gif'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"������",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
case 2
	mess="��ϲ"& sjjh_name &"������˽���㣬�����ʰ��һö�������̫����<IMG src='img/jinbi.gif'>"
	conn.execute "update �û� set ����=����+50000 where ����='"&sjjh_name&"'" 
case 3
	mess="ù��"& sjjh_name &"������˽����ʱ����һ��ɫ��Ů��<font color=red><b>���������һҹ���˻�����,��������100,�����ڳ�����˽������<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ����,״̬ FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����+100,״̬='������' where ����='"&sjjh_name&"'"
	rs.close
case 4
	mess="ù��"& sjjh_name &"������˽����ʱ����һ��ɫ��Ů��<font color=red><b>���������һҹ����÷��,��������100,�����ڳ�����˽������<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ����,״̬ FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����+100,״̬='÷��' where ����='"&sjjh_name&"'"
	rs.close
case 5
	mess="ù��"& sjjh_name &"������˽����ʱ����һ��ɫ��Ů��<font color=red><b>���������һҹ���������,��������100,�����ڳ�����˽������<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ����,״̬ FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����+100,״̬='�����' where ����='"&sjjh_name&"'"
	rs.close
case 6
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶�����������ǿ�ϲ�ɺ�<IMG src='../work/catch/fish.gif'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"�����",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
case 7
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶���С�������ǿ�ϲ�ɺ�<IMG src='../work/catch/fish.gif'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"С����",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
case 8	
	mess="ù��"& sjjh_name &"������˽����ʱ������������ǿ��������5���������ǿɱ�<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ����,״̬ FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����-50000 where ����='"&sjjh_name&"'"
	rs.close	
case 9	
	mess="ù��"& sjjh_name &"������˽����ʱ��������������"& sjjh_name &"����һ������������ʧ5ǧ���ǿɱ�<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ���� FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����-5000 where ����='"&sjjh_name&"'"
	rs.close	
case 10	
	mess="ù��"& sjjh_name &"������˽����ʱ��������������"& sjjh_name &"����һ������������ʧ5ǧ���ǿɱ�<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT ���� FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	conn.execute "update �û� set  ����=����-5000 where ����='"&sjjh_name&"'"
	rs.close
case 11
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶�������ʯ���ǿ�ϲ�ɺ�<IMG src='img/LOOK52.GIF'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"����ʯ",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
case 12
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶����̱�ʯ���ǿ�ϲ�ɺ�<IMG src='img/LOOK52.GIF'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"�̱�ʯ",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
case 13
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶��ź챦ʯ���ǿ�ϲ�ɺ�<IMG src='img/LOOK52.GIF'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"�챦ʯ",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
case 14
	mess="��ϲ"& sjjh_name &"������˽���㣬�����˶���ˮ��ʯ���ǿ�ϲ�ɺ�<IMG src='../work/catch/catch.gif'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"ˮ��ʯ",2)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"					
		
case 15
	mess="��ϲ"& sjjh_name &"������˽���㣡�����˴����㡢����㡢С�����һ����<IMG src='../work/catch/fish.gif'>"
	rs.open "SELECT w6 FROM �û� WHERE ����='"&sjjh_name&"'",conn,3,3
	duyao=add(rs("w6"),"������",1)
	duyao=add(duyao,"�����",1)
	duyao=add(duyao,"С����",1)
	conn.execute "update �û� set  w6='"&duyao&"' where ����='"&sjjh_name&"'"
	rs.close
end select
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>��Ϣ</font>"& sjjh_name &"�������㣺"&mess			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>
<html>
<head>
<META http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../dg/setup.css" rel=stylesheet type="text/css">
<title>��������</title>
<style type=text/css>
body 
 {
	FONT-SIZE: 12px;CURSOR: url(''); COLOR: #f7deac; SCROLLBAR-FACE-COLOR: #f7deac;SCROLLBAR-ARROW-COLOR: #777777;SCROLLBAR-HIGHLIGHT-COLOR: #555555;SCROLLBAR-SHADOW-COLOR: #555555;SCROLLBAR-3DLIGHT-COLOR: #f7deac;SCROLLBAR-TRACK-COLOR: #333333;SCROLLBAR-DARKSHADOW-COLOR: #f7deac;
}
A:hover {
	COLOR: #ff9900;CURSOR: url('');
}
A.menu:hover {
	COLOR: #ccoooo;CURSOR: url('');
}


td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 10pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'f7deac'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff0000;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: 'f7deac'; border: 1 solid 'f7deac'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'f7deac'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'f7deac'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: 'f7deac'; 
	border: f7deacpx solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'f7deac'; color: 'f7deac'; 
	border: 1px solid f7deac;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'f7deac'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'f7deac'}
</style>
</head>

<body bgcolor="#000000" topmargin="0" leftmargin="0">

<table border="0" width="503" height="400" cellspacing="0" cellpadding="0" background="048.gif">
 <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from �û� where ����='" & sjjh_name &"'",conn,3,3
tx=rs("����ͷ��")
tlj=rs("������")
%>
 <tr>
    <td valign="top" height="400">
      <table border="0" width="503" cellspacing="0" cellpadding="0" height="224">
        <tr>
          <td height="224" width="163" valign="top">
            <table border="0" width="165" height="45">
              <tr>
                <td height="41"></td>
              </tr>
            </table>
            <table border="0" width="165" height="102">
              <tr>
                <td width="12" height="98"></td>
                <td width="139" height="98" align="center">&nbsp; <img id=face src="../../ico/<%=tx%>-2.gif" alt=����������� width='70' height='60'></td>  
              </tr>
            </table>
            <table border="0" width="165" height="57">
              <tr>
                <td height="53" width="19"></td>
                <td height="53" width="132" valign="bottom" align="center"><%=sjjh_name%></td>
              </tr>
            </table>
            ��</td>
          <td height="224" width="336" valign="top">
            <table border="0" width="335" height="76">
              <tr>
                <td height="52"></td>
              </tr>
              <tr>
                <td height="16">
                  <div align="center">
                    <center>
                  <table border="1" width="285" bordercolor="#000000" cellspacing="0" cellpadding="0" bordercolordark="#000000" bordercolorlight="#FFFFFF">
                      <tr>
                        <td align="center"><img border="0" src="sailing2.gif" width="240" height="138"></td>
                      </tr>
                    </table>
                    </center>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <div align="center">
        <center>
        <table border="0" width="450" cellspacing="0" height="83">
          <tr>
            <td height="17"></td>
          </tr>
          <tr>
            <td align="center" height="14">��������������</td>
          </tr>
          <tr>
            <td height="14" align="center"><hr color="#FFFFCC" size="1">
</td>
          </tr>
         
          <tr>
            <td height="14">
           <%=mess%></td>
          </tr>
         
          <tr>
            <td height="14" align="center">
             
             ���������ǹŴ�һ��ı�����ܵļ���
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
            <input type=button value=�رմ��� onClick='window.close()' name="button">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
              <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
             
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
          <tr>
            <td height="14" align="center">
             
             </td>
          </tr>
        </table>
        </center>
      </div>
   
      <p>��</td>
  </tr>
</table>


</html>
<script language="JavaScript">                                                
<!--                                                 
var runAs=1;                                                
N = 40;                                                 
Y = new Array();                                                
X = new Array();                                                
S = new Array();                                                
A = new Array();                                                
B = new Array();                                                
M = new Array();                                                
V = (document.layers)?1:0;                                                
iH=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
iW=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
for (i=0; i < N; i++){                                                 
Y[i]=Math.round(Math.random()*iH);                                                
X[i]=Math.round(Math.random()*iW);                                                
S[i]=Math.round(Math.random()*5+2);                                                
A[i]=0;                                                
B[i]=Math.random()*0.1+0.1;                                                
M[i]=Math.round(Math.random()*1+1);                                                
}                                                
if (V){                                                
for (i = 0; i < N; i++)                                                
{document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFF0' CLIP='0,0,"+M[i]+","+M[i]+"'></LAYER>")}                                                
}                                                
else{                                                
document.write('<div style="position:absolute;top:0px;left:0px">');                                                
document.write('<div style="position:relative">');                                                
for (i = 0; i < N; i++)                                                
{document.write('<div id="si" style="position:absolute;top:0;left:0;width:'+M[i]+';height:'+M[i]+';background:#fffff0;font-size:'+M[i]+'"></div>')}                                                
document.write('</div></div>');                                                
}                                                
function snow(){                                                
var H=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
var W=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
var T=(document.layers)?window.pageYOffset:document.body.scrollTop;                                                
var L=(document.layers)?window.pageXOffset:document.body.scrollLeft;                                                
for (i=0; i < N; i++){                                                
sy=S[i]*Math.sin(90*Math.PI/180);                                                
sx=S[i]*Math.cos(A[i]);                                                
Y[i]+=sy;                                                
X[i]+=sx;                                                 
if (Y[i] > H){                                                
Y[i]=-10;                                                
X[i]=Math.round(Math.random()*W);                                                
M[i]=Math.round(Math.random()*1+1);                                                
S[i]=Math.round(Math.random()*5+2);                                                
}                                                
if (V){document.layers['sn'+i].left=X[i];document.layers['sn'+i].top=Y[i]+T}                                                
else{si[i].style.pixelLeft=X[i];si[i].style.pixelTop=Y[i]+T}                                                 
A[i]+=B[i];                                                
}                                                
if(runAs==1){setTimeout('snow()',5);}                                                
}                                                
//-->                                                
</script>                                                
                                                
                                   
<script language=javascript>                                                                                                                     
     function Click(){                                                                                                                    
     alert('E�߽�����ӭ���������治֧�ֻ����Ҽ���');                                                                                                                    
     window.event.returnValue=false;                                                                                                                    
     }                                                                                                                    
     document.oncontextmenu=Click;                                                                                                                    
     </script> 

