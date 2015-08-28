<%'对赌
function gedu1(un,st,mg)
	sql= "select 银两,姓名 FROM 用户 WHERE 姓名='"&un&"' and 银两>=0"
		set rs=conn.execute(sql)
	if rs.EOF or rs.BOF then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你有这么多的钱吗？！');}</script>"
		Response.End 		
	end if
grade=session("yx8_mhjh_usergrade")
xiamoney=abs(int(mg))
if xiamoney<10000 or xiamoney>10000000  then
		Response.Write "<script language=JavaScript>{alert('最少押：10000两，最多1000万！');}</script>"
		Response.End
end if
sql="select 银两,id from 用户 where 姓名='"&un&"'"
set rs=conn.execute(sql)
yin1=rs("银两")
if xiamoney>yin1 then
Response.Write "<script language=JavaScript>{alert('你的身上好象没有这么多钱哦');}</script>"
Response.End 
end if
sql="select 银两 from 用户 where 姓名='"&st&"'"
set rs=conn.execute(sql)
yin2=rs("银两")
if xiamoney>yin2 then
 Response.Write "<script language=JavaScript>{alert('对方身上好象没有这么多钱哦');}</script>"
Response.End
end if
today=now()
abatetime=dateadd("n",-10,today)
abatetimetype="#"&month(abatetime)&"/"&day(abatetime)&"/"&year(abatetime)&" "&hour(abatetime)&":"&minute(abatetime)&":"&second(abatetime)&"#"
sql="delete * from d where c='对' and sj<"&abatetimetype
conn.Execute(sql)
sql="select xm,df from d where xm='"&un&"' and df='"&st&"'"
set rs=conn.execute(sql)
if rs.EOF or rs.BOF then
conn.Execute "insert into d(xm,df,sj,yl,c) values('"&un&"','"&st&"',now(),"&xiamoney&",'对')"
rs.close
gedu1="<b><font color=FF0000>【对赌银两】</font></b>##对%%说：我们来耍一把！赌注为<b><font color=red>"&xiamoney&"</font></b>两，敢吗？<img src='../chatroom/image/sz.gif'><input type=button value='愿意' onclick=""javascript:window.open('geduok.asp?from="&un&"&xiaozhu="&xiamoney&"&to1="&st&"&yn=1','a','width=368,height=122')"" name=""button223""><input type=button value='拒绝' onclick=""javascript:window.open('geduok.asp?from="&un&"&xiaozhu="&xiamoney&"&to1="&st&"&yn=0','a','width=368,height=122')""  name=""button224"">"
else
Response.Write "<script language=JavaScript>{alert('您已经向[" &st&"]提出对赌了，请等对方的回复\n如果对方不作反应，10分钟后请求会自动删除！');}</script>"
Response.End
end if
end function	
%>
