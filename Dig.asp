<!--#include file="Inc/conn.asp"-->
<% 
dim id 
id=request.querystring("id") 
If Request("Page") = "" then 
sql="update pcook_Article set dig = dig where ID= "&ID
conn.execute(sql)
End if
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Sql = "Select * From pcook_Article Where id="&id
	Rs.Open Sql,Conn,3,3
dig=rs("dig")
response.write("document.write('" & dig & "')") 
%>
 function CreateAjax()
 {
    var XMLHttp;
    try
    {
        XMLHttp = new ActiveXObject("Microsoft.XMLHTTP");   
    }
    catch(e)
    {
        try
        {
            XMLHttp = new XMLHttpRequest();     
        }
        catch(e)
        {
            XMLHttp = false;        
        }
    }
    return XMLHttp;     
 }
 
function Dig(id)
{
	_xmlhttp = CreateAjax();
	var url = "<%=SitePath%>save.asp?id="+id+"&n="+Math.random()+"";		
	if(_xmlhttp)    
    {
        var content = document.getElementById("dig"+id);      
		var dig = document.getElementById(id);					
        _xmlhttp.open('GET',url,true);
        _xmlhttp.onreadystatechange=function()
        {
            if(_xmlhttp.readyState == 4)        
            {
                if(_xmlhttp.status == 200)      
                {
                    var ResponseText = unescape(_xmlhttp.responseText);			
					var r=ResponseText.split(",");								
                    if(r[0] == "Dig" )   
                    {
                        alert("���Ѿ�Ͷ��Ʊ�� ������Ͷ�ˣ�");
                        dig.innerHTML=r[1];
                    }
					
                    else if(ResponseText == "NoData")
					{
						alert("��������");	
					}
					else
                    {
						
						dig.innerHTML=ResponseText;
						alert("ͶƱ�ɹ�");
						
						content.innerHTML='�ɹ�';
                    }
                }
                else    
                {
                    alert("���������ش���");
                    top.location.href='<%=SitePath%>index.asp';
                }
            }
            else    
            {
                dig.innerHTML='<img src="<%=SitePath%><%=Skin%>Load.gif">';
            }
        }
        _xmlhttp.send(null);  
    }
    else    
    {
        alert("�����������֧�ֻ�δ���� XMLHttp!");
    }
}