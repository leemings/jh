<%@ LANGUAGE=VBScript codepage ="936" %>
<script lanaguage=javascript>
if(window.name!="aqjh_win"){ var i=1;while (i<=50){window.alert("������ʲôѽ���㵹��ˢ��������������50�Σ���");i=i+1;}top.location.href="../../exit.asp"}
</script>
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
rs.open "select ����,���,�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
hyrmb=rs("����")
taiqiudian=rs("���")
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
<div align="center"><b><font color="#FF0000">��ӭ����<font color=blue><%=session("aqjh_name")%></font>��������ת������</font></b></div>
<div align="center"><b><font color=green size="2"><a href="index.asp">����ת���г�</a></font></b></div>
<br>
        <table width="300" align="center" bordercolor="#FF0000" border="1">
          <tr> 
            <td valign="top" height="8" bgcolor="#000000" > 
              <div align="center">
<%if hyrmb>0 then%>
<p><font size="2"><br>
<%end if%>
 <img border="0" src="jb.gif"> ==&gt;&gt; <img border="0" src="yl.gif"><%if taiqiudian>0 then%></font></p>
  <form method="POST" action="yinzhok.asp" name="aff" onsubmit="bbb.disabled=1">   
  <div align="center"><font color="#FF0000" size="2">��ң�<%=taiqiudian%>�������ת������<%=int(taiqiudian*1000000)%>��   
    </font>  
    <font color="#FFFFFF" size="2"><br>  
    <input type="text" name="input1" size="10" maxlength="5">  
    <input type="submit" value="ת��" name="bbb" class="p9" >    
    </font>   
  </div>   
</form>   
<font color="#FFFFFF" size="2">   
<%else%>   
</font><font color="#0000FF" size="2">   
���Ϊ0����ת��    
</font>   
<font color="#FFFFFF" size="2">   
<br><br>   
<%end if%>   
</font>   
<font color="#FF0000" size="2">   
��<b>��ת��</b>    
</font>   
                </div>  
            </td>  
          </tr>  
          <tr>   
            <td valign="top" bgcolor="#000000" >   
              <p><font color="#FFFF00" size="2">����������Խ���ң����ɽ���������<br>ת����1�����/1000000����!</font></p>  
            </td>  
          </tr>  
        </table>  
<p><font color="#FFFF00" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<img border="0" src="gem01.gif"></font><font size="2" color="#0000FF">���ֽ������� 
2005-01-12</font></p>
</body>  
</html>