<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��Ա��,����,�Ṧ from �û� where ����='"&aqjh_name&"'",conn,2,2
hyrmb=rs("��Ա��")
fl=rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<link rel="stylesheet" href="../../css.css">
<title>����ת��</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#000000">
<div align="center"><b><font color="#FF0000">&nbsp;��ӭ<font color="blue"><%=aqjh_name%></font>���ٷ���ת������&nbsp;</font></b></div>
<div align="center"><b><a href="index.asp"><font color="#FF0000" size="2">����ת���г�</font></a></b></div>
<br>
        <table width="380" align="center" bordercolor="#339966" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" width="370" > 
              <font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <img border="0" src="jb.gif">  
              ====��</font><img border="0" src="fl.gif">
            </td>
          </tr>
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" width="370" > 
              <div align="center">
<%if hyrmb>1 then%>
  <form method="POST" action="jhf.asp" name="af" onsubmit="bbb.disabled=1">
  <div align="center"><font color="#FF0000" size="2">�𿨣�<%=hyrmb%>Ԫ����ת������<%=hyrmb*5000%>��<br>
    <input type="text" name="input" size="10" maxlength="10">
    <input type="submit" value="ת��" name="bbb" class="p9">      
    </font>     
  </div>   
</form>   
<font color="#FF0000" size="2">   
<%else%>   
��Ա��Ϊ1Ԫ����ת��<br><br>  
<%end if%>  
 ����ת��</font>    
                </div> 
            </td> 
          </tr> 
          <tr>  
            <td valign="top" bgcolor="#000000" width="370" >  
              <p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              ��ת����������Ϊ1��50000��</font></p> 
            </td> 
          </tr> 
        </table> 
<p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<img border="0" src="gem01.gif"></font><font size="2" color="#0000FF">���齭������ 
2005-01-12</font></p>
</body> 
</html>