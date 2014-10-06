<%@page import="java.util.HashMap"%>
<%@page import="java.util.Map"%>
<%@page import="java.util.Map"%>
<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Updation Success</title>
<script type="text/javascript" src="js/jquery-2.1.0.min.js"></script>
<link rel="stylesheet" type="text/css" href="css/table.css" />
<link rel="stylesheet" type="text/css" href="css/style.css" />
<script type="text/javascript">
	function detailCompareJSON(apiName)
	{
		var contextpath = '<%=request.getContextPath()%>';
		var myWindow = window.open(contextpath+"/JSONMapperServlet?action=detailCompare&apiName="+apiName,"MsgWindow","width=1400,height=620,scrollbars=1,resizable");
	}
	
	function quickCompareJSON(apiName)
	{
		var contextpath = '<%=request.getContextPath()%>';
		var myWindow = window.open(contextpath+"/JSONMapperServlet?action=quickCompare&apiName="+apiName,"MsgWindow","width=1400,height=620,scrollbars=1,resizable");
	}
</script>
</head>
<body>
	<table border="2" class="tg">
	<tr><th class="tg-header">API Name</th><th class="tg-header" >Status</th><th class="tg-header" colspan="2">Validate</th></tr> 
	<%
		int success = 0,failed = 0;
		String action =(String) request.getAttribute("action");
		HashMap<String,String> resultConfig = (HashMap<String,String>)request.getAttribute("result");
	   for(Map.Entry<String, String> result : resultConfig.entrySet()){
		   %>
		   <tr>
		   <td class="tg-col-odd"><%= result.getKey().toString()%></td>
		   <%if(result.getValue().toString().contains("Sorry")){
			   failed++;
			   %>
			   <td bgcolor="#D76A81"><%=result.getValue().toString()%></td>
			   <td><input type="submit" class="add-photo-btn" id="detailCompare" value="Detail Compare" disabled onclick="return detailCompareJSON('<%=result.getKey()%>');"></td>
		   		<td><input type="submit" class="add-photo-btn" id="quickCompare" value="Quick Compare" disabled onclick="return quickCompareJSON('<%=result.getKey()%>');"></td>
		   <% }else{ success++; %>
		    <td class="tg-col-even"><%=result.getValue().toString()%></td>
		   	<td><input type="submit" class="add-photo-btn" id="detailCompare" value="Detail Compare"  onclick="return detailCompareJSON('<%=result.getKey()%>');"></td>
		   	<td><input type="submit" class="add-photo-btn" id="quickCompare" value="Quick Compare"  onclick="return quickCompareJSON('<%=result.getKey()%>');"></td>
		   <%} %>
		   </tr>
	  <% }	%>
	<tr><td align="center" colspan="2">Success :<%=success%></td><td align="center" colspan="2">Failed :<%=failed%></td></tr> 
	</table>
	</div>
</body>
</html>