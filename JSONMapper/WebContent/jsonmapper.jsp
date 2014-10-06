<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- ResourceBundle resource = ResourceBundle.getBundle("com.att.blackflag.json.resource.ApplicationProperties",Locale.getDefault()); %>--> 

<%@page import="java.util.ResourceBundle" %>
<%@page import="java.util.HashMap" %>
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
	
	function addNewEnvironment(){
		var errors = document.getElementById("errors");
		errors.innerHTML = '';
		var envNameObj = document.getElementById("envName");

		var success = true;
		if (envNameObj != null && envNameObj.value == -1) {
			errors.innerHTML = "Please Select Environment Name";
			success = false;
		} 
		if(success) {
		var contextpath = '<%=request.getContextPath()%>';
		document.envForm.action = contextpath+'/JSONMapperServlet?action=addJSON';
		} else {
			return false;
		}
	}
	
	function updateEnvironment()
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
	<h1>BlagFlag Environment Configuration</h1>
	<br><br><br>
	<p style="color: red;font-family: sans-serif ;font-size: 12px" id="errors"></p>
	<div  class="CSSTableGenerator">
	<table>
			<tr>
				<td>Environment Name:</td>
				<td colspan="2">
				<div  class="styled-select">
					 <select id="envName" name="envName">
					   <option value="-1">--Please Select--</option>
					   <% String[] envList = resource.getString("ENVIRONMENT-NAMES").split(",");
					   		for(String env : envList){ %>
					   			<option value='<%=env%>'><%=env%></option>		
					   		<%}
					   %>
					</select>
					</div>
				</td>
			</tr>
			<tr>
				<td>Excel Path:</td>
				<td>
					<p class="form">
					<input type="text" id="path" name="path" size="45" value=""/>
					<label class="add-photo-btn" >upload<span><input type="file" id="excelFileName" name="excelFileName" onfocus="document.getElementById('path').value = document.getElementById('excelFileName').value"></span></label>
					</p>
				</td>
			</tr>
			<tr>
				<div align="center">
				<td colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="submit" class="button_example" name="updateJSONbtn" value="Update" onclick="return updateEnvironment();">&nbsp;&nbsp;&nbsp;<input type="submit" class="button_example" name="addJSONbtn" value="Add" onclick="return addNewEnvironment();"></td>
				</div>
			</tr>
			<%
				String error = (String)request.getAttribute("error");
				if(error!=null){%>
							<tr>
					<td colspan ="3">Please select valid excel configuration file</td>
								</tr>
			<%	} %>

	</table>
	</div>
	</center>
	</form>
</body>
</html>