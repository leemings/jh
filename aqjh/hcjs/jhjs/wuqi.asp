<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<HTML><HEAD><TITLE>������</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE>TD {
	FONT-SIZE: 9pt
}
</STYLE>

<SCRIPT language=javascript>
function showwp(ename,cname,tx,tj,lb,liliang,fangyu,sudu,price){
	zl.innerHTML='<input type=hidden name=wpname value='+cname+'><img border="0" src="images/'+ename+'" width="35" height="35">&nbsp;'+cname+'[<font color=blue>'+lb+'</font>]<br><br>������'+liliang+'<br>������'+fangyu+'<br>�;ã�'+sudu+'<br>����:'+tx+'<br>����:'+tj+'<br>�۸�'+price+'Ԫ'
}
function yesno(name) {
if(confirm("�����ڽ�Ҫ������� ["+name+"]  \n�����ڿ��Ե�ȷ������")) {document.form1.submit();}
}
</SCRIPT>

<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=5>
<DIV align=center>
<CENTER>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width="621">
      <FORM action=wuqiok.asp method=post name=form1>
        <td width="4">&nbsp;
        <TR bgcolor="#307028"> 
          <TD colSpan=5><marquee border='1' onmouseover=this.stop(); onmouseout=this.start();>��ӭ������[���齭��]�����ꡭ��</marquee></TD>
        </TR>
        <TR> 
          <TD bgColor=#307028 rowSpan=5 width="4"><IMG height=379 src="images/index_02.gif" width=4></TD>
          <TD bgColor=#307028 height=379 rowSpan=5 vAlign=top width=405> 
            <TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
 <TBODY> 
 <TR> 
 <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b where b='����' or b='װ��' or b='����' or b='˫��' or b='ͷ��' or b='����' order by b,h",conn,3,3
do while not rs.bof and not rs.eof
tj=replace(rs("l"),chr(39),"")
Response.Write "<TD  width='16%'><A href='#' onclick="&chr(34)&"showwp('"&rs("i")&"','"&rs("a")&"','"&rs("k")&"','"&tj&"','"&rs("b")&"',"&rs("f")&","&rs("g")&","&rs("j")&","&rs("h")&");"&chr(34)&" title='"&rs("c")&"'><IMG border=0 height=35 src='images/"&rs("i")&"' width=35></A></TD>"
x=x+1
if x/6=int(x/6) then Response.Write "</tr><tr>"
rs.movenext
loop
Response.Write "</tr>"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

 </TBODY> 
 </TABLE>
  </TD>
  <CENTER>
            <TD bgColor=#408c30 rowSpan=5 width="5"><IMG height=420 
      src="images/index_04.gif" width=5></TD>
            <TD width="203" height="101"><IMG height=145 src="images/index_05.gif" width=202></TD>
            <TD rowSpan=5 width="5"><IMG height=420 src="images/index_04.gif" 
  width=5></TD>
   </center>
        </TR>
        <TR> 
          <TD border="0" width="203" background="images/index_09.gif">&nbsp;</TD>
        </TR>
        <TR> 
          <TD bgColor=#707470 height=136 id=zl vAlign=top width=203> <input type=hidden name=wpname value=''>
            <P style="LINE-HEIGHT: 150%"><BR>
 &nbsp; �ݺ�����������Ϊ���ṩս��������װ��,��ӭ��λ����ѡ����</P>
          </TD>
        </TR>
        <TR> 
          <TD id=esoteric width="203" background="images/index_09.gif">&nbsp;</TD>
        </TR>
        <TR> 
          <TD bgColor=#707470 height=111 vAlign=top width=203> 
            <P align=center><BR>
 ��ȷ��������������<BR>
 <BR>
 <INPUT height=34  onclick="if(document.form1.wpname.value!=''){yesno(document.form1.wpname.value);return false;}else{return false;}" 
      src="images/queding.gif" type=image value=�ύ width=98>
 </P>
          </TD>
        </TR>
      </FORM>
    </TABLE>
  </CENTER></DIV></BODY></HTML>