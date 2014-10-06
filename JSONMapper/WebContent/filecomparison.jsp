<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%@page import="java.util.Iterator"%>
<%@page import="java.util.HashMap"%>
<%@page import="java.util.List"%>
<%@page import="java.util.Map"%>
<%@page import="java.util.Set"%>
<%@page import="java.util.LinkedHashMap"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="text/javascript" src="js/jquery-2.1.0.min.js"></script>
<link rel="stylesheet" type="text/css" href="css/table.css" />
<link rel="stylesheet" type="text/css" href="css/style.css" />

	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>Compare JSON Conguration changes</title>
	
</head>
<body background="images/imagesbg2.png">
	<table border="0" class="codecompare">
			<tr bgcolor="#585858">
				<td style="width:2%;color:#EFF5FB">L.No</td>
				<td style="text-align:center;color:#EFF5FB">Old File</td>
				<td style="text-align:center;color:#EFF5FB">New File</td>
			</tr>
					<%
					LinkedHashMap<String, LinkedHashMap<String,String>> dispMap = (LinkedHashMap<String, LinkedHashMap<String,String>>) request.getAttribute("diffResult");
					
					if(dispMap != null)
					{
						Iterator itr = dispMap.entrySet().iterator();  
				       
				        while(itr.hasNext())
				        {
				        	LinkedHashMap<String,String> diffMap = new LinkedHashMap<String,String>();
				        	
				        	Map.Entry entry = (Map.Entry) itr.next();
				        	diffMap = (LinkedHashMap<String,String>)entry.getValue();
				        	for(Map.Entry<String, String> diff : diffMap.entrySet()){
				        	%>
				        		<tr><%=entry.getKey() %><%=diff.getKey() %><%=diff.getValue() %></tr>
				        		<%
			        		}
				        }
					}
					%>
	</table>
</body>
</html>