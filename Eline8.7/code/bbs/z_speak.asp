<%if isnull(speak) then
	dim speak
	speak=""
end if%>
<br>
发帖语气 <SELECT name="speak">
<OPTION value="说道" <%if isnull(speak) or speak="" then%>selected<%end if%>>选择语气</OPTION> 
<OPTION value="说道" <%if speak="说道" then%>selected<%end if%>>说道</OPTION>
<OPTION value="骂道" <%if speak="骂道" then%>selected<%end if%>>骂道</OPTION>
<OPTION value="笑着说" <%if speak="笑着说" then%>selected<%end if%>>笑着说</OPTION> 
<OPTION value="哭着说" <%if speak="哭着说" then%>selected<%end if%>>哭着说</OPTION> 
<OPTION value="啜泣道" <%if speak="啜泣道" then%>selected<%end if%>>啜泣道</OPTION>
<OPTION value="仰头说" <%if speak="仰头说" then%>selected<%end if%>>仰头说</OPTION>
<OPTION value="咆哮道" <%if speak="咆哮道" then%>selected<%end if%>>咆哮道</OPTION>
<OPTION value="表扬道" <%if speak="表扬道" then%>selected<%end if%>>表扬道</OPTION>
<OPTION value="批评道" <%if speak="批评道" then%>selected<%end if%>>批评道</OPTION>
<OPTION value="微笑着说" <%if speak="微笑着说" then%>selected<%end if%>>微笑着说</OPTION> 
<OPTION value="急忙问道" <%if speak="急忙问道" then%>selected<%end if%>>急忙问道</OPTION>
<OPTION value="弯下腰说" <%if speak="弯下腰说" then%>selected<%end if%>>弯下腰说</OPTION>
<OPTION value="感动地说" <%if speak="感动地说" then%>selected<%end if%>>感动地说</OPTION>
<OPTION value="激动地说" <%if speak="激动地说" then%>selected<%end if%>>激动地说</OPTION>
<OPTION value="疑惑地说" <%if speak="疑惑地说" then%>selected<%end if%>>疑惑地说</OPTION>
<OPTION value="不屑地说" <%if speak="不屑地说" then%>selected<%end if%>>不屑地说</OPTION>
<OPTION value="难过地说" <%if speak="难过地说" then%>selected<%end if%>>难过地说</OPTION>
<OPTION value="狂笑着说道" <%if speak="狂笑着说道" then%>selected<%end if%>>狂笑着说道</OPTION> 
<OPTION value="冷笑着说道" <%if speak="冷笑着说道" then%>selected<%end if%>>冷笑着说道</OPTION> 
<OPTION value="笑骂着说道" <%if speak="笑骂着说道" then%>selected<%end if%>>笑骂着说道</OPTION>
<OPTION value="濠淘大哭道" <%if speak="濠淘大哭道" then%>selected<%end if%>>濠淘大哭道</OPTION> 
<OPTION value="哭着哀求道" <%if speak="哭着哀求道" then%>selected<%end if%>>哭着哀求道</OPTION> 
<OPTION value="跪地求饶道" <%if speak="跪地求饶道" then%>selected<%end if%>>跪地求饶道</OPTION> 
<OPTION value="深情地说道" <%if speak="深情地说道" then%>selected<%end if%>>深情地说道</OPTION>
<OPTION value="热情地说道" <%if speak="热情地说道" then%>selected<%end if%>>热情地说道</OPTION>
<OPTION value="浪漫地说道" <%if speak="浪漫地说道" then%>selected<%end if%>>浪漫地说道</OPTION>
<OPTION value="天真地说道" <%if speak="天真地说道" then%>selected<%end if%>>天真地说道</OPTION>
<OPTION value="鞠了一躬说" <%if speak="鞠了一躬说" then%>selected<%end if%>>鞠了一躬说</OPTION>
<OPTION value="挥了挥手说" <%if speak="挥了挥手说" then%>selected<%end if%>>挥了挥手说</OPTION>
<OPTION value="气愤地说道" <%if speak="气愤地说道" then%>selected<%end if%>>气愤地说道</OPTION>
<OPTION value="仰望苍天道" <%if speak="仰望苍天道" then%>selected<%end if%>>仰望苍天道</OPTION>
<OPTION value="大声地说道" <%if speak="大声地说道" then%>selected<%end if%>>大声地说道</OPTION>
<OPTION value="小声地说道" <%if speak="小声地说道" then%>selected<%end if%>>小声地说道</OPTION>
<OPTION value="严肃地说道" <%if speak="严肃地说道" then%>selected<%end if%>>严肃地说道</OPTION>
<OPTION value="点头称赞道" <%if speak="点头称赞道" then%>selected<%end if%>>点头称赞道</OPTION>
<OPTION value="很欣赏地说" <%if speak="很欣赏地说" then%>selected<%end if%>>很欣赏地说</OPTION>
<OPTION value="啧啧嘴说道" <%if speak="啧啧嘴说道" then%>selected<%end if%>>啧啧嘴说道</OPTION>
<OPTION value="悲伤地说道" <%if speak="悲伤地说道" then%>selected<%end if%>>悲伤地说道</OPTION>
<OPTION value="眼泪嗒叭地说" <%if speak="眼泪嗒叭地说" then%>selected<%end if%>>眼泪嗒叭地说</OPTION> 
<OPTION value="提高嗓门说道" <%if speak="提高嗓门说道" then%>selected<%end if%>>提高嗓门说道</OPTION>
<OPTION value="降低声音说道" <%if speak="降低声音说道" then%>selected<%end if%>>降低声音说道</OPTION>
<OPTION value="背过身去说道" <%if speak="背过身去说道" then%>selected<%end if%>>背过身去说道</OPTION>
<OPTION value="竖着大姆指说" <%if speak="竖着大姆指说" then%>selected<%end if%>>竖着大姆指说</OPTION>
<OPTION value="头也不抬地说" <%if speak="头也不抬地说" then%>selected<%end if%>>头也不抬地说</OPTION>
<OPTION value="莫名其妙地说" <%if speak="莫名其妙地说" then%>selected<%end if%>>莫名其妙地说</OPTION>
<OPTION value="义愤填膺地说" <%if speak="义愤填膺地说" then%>selected<%end if%>>义愤填膺地说</OPTION>
<OPTION value="瞥了一眼说道" <%if speak="瞥了一眼说道" then%>selected<%end if%>>瞥了一眼说道</OPTION>
<OPTION value="幸灾乐祸地说" <%if speak="幸灾乐祸地说" then%>selected<%end if%>>幸灾乐祸地说</OPTION>
<OPTION value="神秘兮兮地说" <%if speak="神秘兮兮地说" then%>selected<%end if%>>神秘兮兮地说</OPTION>
<OPTION value="可怜兮兮地说" <%if speak="可怜兮兮地说" then%>selected<%end if%>>可怜兮兮地说</OPTION>
<OPTION value="卖着关子说道" <%if speak="卖着关子说道" then%>selected<%end if%>>卖着关子说道</OPTION>
<OPTION value="红着脸害羞地说" <%if speak="红着脸害羞地说" then%>selected<%end if%>>红着脸害羞地说</OPTION>
<OPTION value="迫不急待地问道" <%if speak="迫不急待地问道" then%>selected<%end if%>>迫不急待地问道</OPTION>
<OPTION value="浑身发抖地说道" <%if speak="浑身发抖地说道" then%>selected<%end if%>>浑身发抖地说道</OPTION>
<OPTION value="抓了抓头皮说道" <%if speak="抓了抓头皮说道" then%>selected<%end if%>>抓了抓头皮说道</OPTION>
<OPTION value="痛哭流涕地说道" <%if speak="痛哭流涕地说道" then%>selected<%end if%>>痛哭流涕地说道</OPTION> 
<OPTION value="挥着拳头威胁道" <%if speak="挥着拳头威胁道" then%>selected<%end if%>>挥着拳头威胁道</OPTION>
<OPTION value="一本正经地说道" <%if speak="一本正经地说道" then%>selected<%end if%>>一本正经地说道</OPTION>
<OPTION value="把脸凑上去说道" <%if speak="把脸凑上去说道" then%>selected<%end if%>>把脸凑上去说道</OPTION>
<OPTION value="目不转睛地说道" <%if speak="目不转睛地说道" then%>selected<%end if%>>目不转睛地说道</OPTION>
<OPTION value="瞪大眼睛吃惊地说" <%if speak="瞪大眼睛吃惊地说" then%>selected<%end if%>>瞪大眼睛吃惊地说</OPTION>
<OPTION value="上气不接下气地说道" <%if speak="上气不接下气地说道" then%>selected<%end if%>>上气不接下气地说道</OPTION>
<OPTION value="耸了耸肩无奈地说道" <%if speak="耸了耸肩无奈地说道" then%>selected<%end if%>>耸了耸肩无奈地说道</OPTION>
<OPTION value="叹了口气失望地说道" <%if speak="叹了口气失望地说道" then%>selected<%end if%>>叹了口气失望地说道</OPTION>
<OPTION value="嘟着嘴不满意地说道" <%if speak="嘟着嘴不满意地说道" then%>selected<%end if%>>嘟着嘴不满意地说道</OPTION>
<OPTION value="打破沙锅问到底地说道" <%if speak="骂道" then%>selected<%end if%>>打破沙锅问到底地说道</OPTION>
<OPTION value="舔了舔流出来的口水说" <%if speak="骂道" then%>selected<%end if%>>舔了舔流出来的口水说</OPTION>
</SELECT>