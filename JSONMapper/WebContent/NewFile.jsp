<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%@ page import = "java.util.ResourceBundle" %>
<% ResourceBundle resource = ResourceBundle.getBundle("com.att.blackflag.json.resource.ApplicationProperties"); %> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="text/javascript" src="js/jquery-2.1.0.min.js"></script>
<link rel="stylesheet" type="text/css" href="css/table.css" />
<link rel="stylesheet" type="text/css" href="css/style.css" />

	<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
	<title>Home</title>
	<script type="text/javascript">
	function addEnv()
	{
		var contextpath = '<%=request.getContextPath()%>';
		document.envForm.action = contextpath+'/JSONMapperServlet?action=addEnvironment';
	}
	
	function modifyEnv()
	{
		var contextpath = '<%=request.getContextPath()%>';
		document.envForm.action = contextpath+'/JSONMapperServlet?action=modifyEnvironment';
	}
	
	function updateJSON()
	{
		var errors = document.getElementById("errors");
		errors.innerHTML = '';
		var envNameObj = document.getElementById("envName");
		//var apiNameObj = document.getElementById("apiName");
		var excelFileNameObj = document.getElementById("excelFileName");
		//var jsonNameObj = document.getElementById("jsonName");
		var success = true;
		if (envNameObj != null && envNameObj.value == -1) {
			errors.innerHTML = "Please Select Environment Name";
			success = false;
		} else if (excelFileNameObj != null && excelFileNameObj.value == '') {
			errors.innerHTML = "Please browse Excel File";
			success = false;
		}
		if(success) {
		var contextpath = '<%=request.getContextPath()%>';
		document.envForm.action = contextpath+'/JSONMapperServlet?action=updateJSON';
		} else {
			return false;
		}
	}
	
	function compareJSON()
	{
		var contextpath = '<%=request.getContextPath()%>';
		document.envForm.action = contextpath+'/JSONMapperServlet?action=compareJSON';
	}
	
	</script>
</head>
<body background="images/chatback.png">

	<form name="envForm" action="JSONMapperServlet" method="post">
	<center><br><br>
	<h1>BlagFlag Environment Configuration Files Comparision</h1>
	<br><br><br>
	<p style="color: red;font-family: sans-serif ;font-size: 12px" id="errors"></p>
	<!-- <div  class="CSSTableGenerator">-->
	<table border=1>
			<tr>
				<td> Old Config.json</td>
				<td>Updated Config.json</td>
			</tr>
			<tr>
			<td>
					 Hi I am Old Json file for comparision. I have<br> 
					 to modify lot of places, so------------------<br>
					 plz do the need{} ''''''''''''''''''''''ful..
					 "xpath": "/TargetEndpoint/HTTPTargetConnection/Properties/Property[@name\u003d\u0027io.timeout.millis\u0027]",
				</td>
				<td>
					 Hi I am Updated Json file for comparision
				</td>
			</tr>
			
</table>
	<!-- </div>-->
	</center>
	</form>
</body>
</html>