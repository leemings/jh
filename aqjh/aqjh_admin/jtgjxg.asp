<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �� where id="& id,conn
l=rs("����")
a1=""
a2=""
a3=""
a4=""
a5=""
a6=""
a7=""
a8=""
select case l
	case "�ɻ�"
		a1="selected"
	case "�γ�"
		a2="selected"
	case "Ħ�г�"
		a3="selected"
	case "��"
		a4="selected"
	case "���г�"
		a5="selected"
	case "����"
		a6="selected"
end select
%>
<html>
<head>
<title>��ͨ�����޸�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../JHIMG/bk_Hc3w.gif">
<p align="center">�޸Ľ�ͨ����</p>
<form name="form1" method="post" action="jtgjxgok.asp?id=<%=rs("id")%>">
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    ���ͣ�    
    <select name="lx">
      <option value="�ɻ�" <%=a1%>>�ɻ�</option>
      <option value="�γ�" <%=a2%>>�γ�</option>
      <option value="Ħ�г�" <%=a3%>>Ħ�г�</option>
      <option value="��" <%=a4%>>��</option>
      <option value="���г�" <%=a5%>>���г�</option>
      <option value="����" <%=a6%>>����</option>
    </select>
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  ͼƬ��     
    <input type="text" name="tp" value="<%=rs("ͼƬ")%>">
  <font color="#FF0000">ע���˴�ͼƬ��Ҫ��·��</font></font></p>   
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  �۸�     
    <input type="text" name="jg" value="<%=rs("�۸�")%>">
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
  ս���ȼ���     
    <input type="text" name="zddj" value="<%=rs("ս���ȼ�")%>" size="20">
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  &nbsp;&nbsp;&nbsp; ����ȼ��� </font><font size="2"><input type="text" name="gldj" value="<%=rs("����ȼ�")%>" size="20">  
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  ��Աר�� ��     
    <%if rs("��Աר��")=true then%>
    <select name="hyzy">    
      <option value="true" selected>��</option>
      <option value="false">��</option>
    </select>
    <%else%>    
    <select name="hyzy">    
      <option value="true" >��</option>
      <option value="false" selected>��</option>
    </select>
    <%end if%>    
    </font></p>
  <p align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
  ˵����     
    <textarea name="sm" rows="4" cols="56"><%=rs("˵��")%></textarea>
    </font></p>
  <p align="center"><font size="2">˵�����ͼƬλ��Ҫ��׼ȷ��·��</font></p>
  <p align="center"> <font size="2"> 
    <input type="submit" name="Submit" value="�޸�">
    <input type="reset" name="��д" value="����">
    </font></p>
</form>
<font size="2"> 
<%rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</font> 
<p align="center"><font size="2"><a href="manjtgj.asp?lx=<%=l%>">�� ��</a> </font></p>
</body>
</html>
