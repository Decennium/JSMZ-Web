<br clear=both>
<form method="POST" action="result.asp">
<div class="footer">
<input type="text" name="q" size="31" maxlength="250" value="<%=request("q")%>" title="ËÑË÷" onfocus="this.select()" onmouseover="this.select()">
<input type="submit" name="btnS" value="ËÑË÷">
<input type="hidden" name="page" value="1">
<input type="hidden" name="newwindow" value=1>
<input type="hidden" name="stype" value="<%=request("stype")%>">

</div>
</form>
<!--#include file="../include/bottom.asp"-->
