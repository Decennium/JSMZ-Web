function choose_date_czw(date_id,objtd){
if(date_id=="choose_date_czw_close"){
    document.getElementById("choose_date_czw_id").style.display="none";
    return;
}
if(objtd!=undefined){
	if(objtd=="choose_date_czw_now"){
		var currentTime = new Date();
		var year1 = currentTime.getFullYear();
		var month1 = currentTime.getMonth() + 1;
		var day1 = currentTime.getDate();

		document.getElementById(date_id).value = year1 + "-" + month1 + "-" + day1;
		document.getElementById("choose_date_czw_id").style.display="none";
		return;
	}else{
		var year1 = document.getElementById("choose_date_czw_year").value;
		var month1 = document.getElementById("choose_date_czw_month").value;
		document.getElementById(date_id).value=year1+"-"+month1+"-"+objtd.innerHTML;
	}
	document.getElementById("choose_date_czw_id").style.display="none";
	return;
}
var nstr=new Date(); //��ǰ
if(document.getElementById("choose_date_czw_year")!=null){
    var year = document.getElementById("choose_date_czw_year").value;
    var month = document.getElementById("choose_date_czw_month").value;
    var str=year+"/"+month+"/1";
    nstr=new Date(str); //��ǰ
}
var ynow=nstr.getFullYear(); //���
var mnow=nstr.getMonth(); //�·�
var dnow=nstr.getDate(); //��������
var n1str=new Date(ynow,mnow,1); //���µ�һ��
var firstday=n1str.getDay(); //���µ�һ�����ڼ�
function is_leap(year) {
   return (year%100==0 ? res=(year%400==0 ? 1 : 0) : res=(year%4==0 ? 1: 0));
}
var dstr="<select id=\"choose_date_czw_year\" onchange=\"choose_date_czw('"+date_id+"')\">";
for(var y=1901;y<2050;y++){
    if(y==ynow){
        dstr+="<option value='"+y+"' selected>"+y+"</option>"
    }else{
        dstr+="<option value='"+y+"'>"+y+"</option>"
    }
}
dstr+="</select><select id=\"choose_date_czw_month\" onchange=\"choose_date_czw('"+date_id+"')\">";
for(var m=1;m<13;m++){
    if(parseInt(mnow+1)==m){
        dstr+="<option value='"+m+"' selected>"+m+"</option>"
    }else{
        dstr+="<option value='"+m+"'>"+m+"</option>"
    }
}
dstr+="</select>"
dstr+="<span style='padding:2px;margin:2px;background-color:#E5ECF9;cursor:pointer;' onclick=\"choose_date_czw('"+date_id+"','choose_date_czw_now')\">����</span>"
dstr+="<span style='padding:2px;margin:2px;background-color:#C0C0C0;cursor:pointer;' onclick=\"choose_date_czw('choose_date_czw_close')\">�ر�</span>";
//һ�����߰�ʮ��(ʮ����),��ʮһ��������;�����Ŷ�(ʮһ��)��ʮ��,Ψ�ж��¶�ʮ��(�����ʮ��).
var m_days = new Array(31,28+is_leap(ynow),31,30,31,30,31,31,30,31,30,31);
var tr_str=Math.ceil((m_days[mnow] + firstday)/7);
dstr+="<table border='0' cellpadding='5' cellspacing='0'><tr bgcolor='#C0C0C0'><td>��</td><td>һ</td><td>��</td><td>��</td><td>��</td><td>��</td><td>��</td></tr>";
var dqdate=new Date(); //��ǰ
for(i=0;i<tr_str;i++) { //���for��� - tr��ǩ
   dstr+="<tr>";
   for(k=0;k<7;k++) { //�ڲ�for��� - td��ǩ
      idx=i*7+k; //���Ԫ����Ȼ���
      date_str=idx-firstday+1; //��������
     if(date_str<=0 || date_str>m_days[mnow]){
          dstr+="<td>&nbsp;</td>";
     }else{
        if(ynow==dqdate.getFullYear() && mnow==dqdate.getMonth() && dqdate.getDate()==date_str){
            dstr+="<td onmouseover=\"this.style.backgroundColor='#E5ECF9'\" onmouseout=\"this.style.backgroundColor='#E5ECF9'\" onclick=\"choose_date_czw('"+date_id+"',this)\" style='cursor:pointer; background-color:#E5ECF9;'>"+date_str+"</td>";
        }else{
            dstr+="<td onmouseover=\"this.style.backgroundColor='#E5ECF9'\" onmouseout=\"this.style.backgroundColor='#fff'\" onclick=\"choose_date_czw('"+date_id+"',this)\" style='cursor:pointer;'>"+date_str+"</td>";
        }
     }
   }
   dstr+="</tr>";
}
dstr+="</table>";
if(document.getElementById("choose_date_czw_id")==null){
var obj = document.getElementById(date_id);
var odiv = document.createElement("div");
odiv.id="choose_date_czw_id";
odiv.innerHTML=dstr;
odiv.style.position="absolute";
odiv.style.border="1px #0CF solid";
odiv.style.fontSize="12px";
odiv.style.backgroundColor="white";
odiv.style.zIndex=99999;
odiv.style.top=getTopY(obj)+"px";
odiv.style.left=getTopX(obj)+"px";
//obj.onblur="document.getElementById(\"choose_date_czw_id\").style.display=\"none\";"
document.body.appendChild(odiv);
}else{
    document.getElementById("choose_date_czw_id").style.display="block";
    document.getElementById("choose_date_czw_id").innerHTML=dstr;
}
}

function getTopX(elem)
{
	return elem.offsetParent?(elem.offsetLeft+getTopX(elem.offsetParent)):elem.offsetLeft;
}

function getTopY(elem)
{
	return elem.offsetParent?(elem.offsetTop+getTopY(elem.offsetParent)):elem.offsetTop;
}
