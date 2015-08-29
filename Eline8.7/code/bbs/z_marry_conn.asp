<%
dim connj,name,name3,id,rs1,txt_sMateNick,user,message,sex1,sex2
dim giftname,price,imageurl,nowmoney,splosh,rs3,rs4,ok,str1,ount1
'dim agree,toname,content,guige,theyear,them,thed,theh,lasttime
dim agree,toname,guige,theyear,them,thed,theh,lasttime
dim conn1str,fumoney,total,yourmoney,name2,money,thename1,thename2
dim db2
db2="data/e_marry.asp"
set connj = Server.CreateObject("ADODB.connection")
conn1str="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db2)
connj.Open conn1str
	
sub marry_nav()
	response.write "<table border=0 cellpadding=3 cellspacing=1 align=center class=tableborder1>"&_
					"<tr>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_index.asp>婚介中心</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_dglistall.asp>点歌祝福</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_courtship.asp>求婚中心</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_statute.asp>婚姻法规</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_manager.asp>月老版块</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_register.asp>结婚登记</a></td>"&_
					"<td class=TopLighNav1 align=center><a href=z_marry_cabaret.asp>社区酒店</a></td>"&_
					"</tr></table><br>"
end sub%>