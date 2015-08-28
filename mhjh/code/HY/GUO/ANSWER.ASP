<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if Session("usepro")<>"1" then
Session("usepro")=""
response.redirect "index.asp"
end if
if session("yx8_mhjh_username")="" then Response.Redirect "../error.asp?id=016"
   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & _
             Server.MapPath("qwas.asp")
   Set rs = Server.CreateObject("ADODB.Recordset")
   sql = "Select * from QUESTION "
   rs.Open sql, conn, adOpenStatic
%>
<script language="javascript">
function popwin(id)
{		
window.open("edit.asp?id="+id+"",'','menubar=yes,scrollbars=yes resizable=yes width=480,height=300');
}
</script>
<script language="javascript">
function click() {
if (event.button==2) {
alert('PPMM盯着你，象要杀人：不要作弊哟~！！！')
}
}
document.onmousedown=click
</script>
<html>
<head>
<title>问题考验</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet>
 </head>

<body oncontextmenu=self.event.returnValue=false background="../../chatroom/bg1.gif" >
<form method="post" action="answer2.asp" name="form">
 <p align="center">还剩下
 <input type="text" name="remain_time" disabled size="2" value="80" maxlength="2" style="width:20;height:16;border:0 solid #FFFFFF; font-size:10.8pt; ">秒        

  <table width="96%" cellspacing="1" cellpadding="0" border="1" bordercolor="#000000" align="center">
    <tr> 

      <td width="18%" height="17"></td>
    </tr>
    <tr> 
      <td width="82%" height="17">
        <p style="line-height: 105%; margin-left: 6; margin-right: 6"><b>题目</b></p>
      </td>
      <td width="18%" height="17">&nbsp;</td>
    </tr>
    <%dim i,ran
			for i=1 to 10
			randomize
			ran=(int((6-0+1)*rnd+0))&(int((9-0+1)*rnd+0))&(int((9-0+1)*rnd+0))
			ran=clng(ran)
			rs.movefirst
			rs.move ran-1
			If rs.EOF Then Exit For
			%> 
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">第<%=I%>题：<%=rs("question")%></p></td>
      <td width="18%">选择答案 
        <input type="hidden" name="ann<%=i%>" value="<%=rs("answer")%>">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">A：<%=rs("A")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="A">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">B：<%=rs("b")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="B">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">C：<%=rs("c")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="C">
      </td>
    </tr>
    <tr> 
      <td width="82%" height="26"><p style="line-height: 105%; margin-left: 6; margin-right: 6">D：<%=rs("d")%></P></td>
      <td width="18%" height="26"> 
        <input type="radio" name="AN<%=I%>" value="D">
      </td>
    </tr>
    <%
	     Next
	    %> 
  </table>

  <div align="center"> 
    <input type="submit" name="Submit" value="提交">

  </div>

</form>
</html>              <script language="javascript">        
      
function time_step()        
{        
        with(document.form.remain_time)        
        {        
                value=(value!=0)?value-1:0;        
        
                if(document.form.remain_time.value<=0)        
                {        
                        alert("您已经超过时间，知识问答结束！");        
						location.href ='INDEX.ASP'        
				}        
				else        
				{        
			       		window.setTimeout("time_step()",1000);        
				}        
		}        
}        
time_step();        
</script>         
