package com.att.blackflag.json.servlet;

import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.att.blackflag.json.helper.JSONObjectMapper;
import com.att.blackflag.json.util.JSONComparator;

/**
 * Servlet implementation class JSONMapperServlet
 */
public class JSONMapperServlet extends HttpServlet {
	
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public JSONMapperServlet() {
        super();
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doPost(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String action = request.getParameter("action");
		String envName = request.getParameter("envName");
		String excelFileName = request.getParameter("excelFileName");
		RequestDispatcher dispatcher = null;
		try{
			if(action.equalsIgnoreCase("addJSON")) {
				JSONObjectMapper jom=new JSONObjectMapper(envName, excelFileName);
				HashMap<String, String> result = jom.addNewEnvConfig();
				request.setAttribute("result", result);
				request.setAttribute("action", "add");
	            dispatcher = getServletContext().getRequestDispatcher("/mapperresult.jsp");
			}else if(action.equalsIgnoreCase("updateJSON")) {
				if(excelFileName!=null && !excelFileName.endsWith("xlsx") && !excelFileName.endsWith("xls")){
					request.setAttribute("error", "error");
		            dispatcher = getServletContext().getRequestDispatcher("/jsonmapper.jsp");
				}else{
					JSONObjectMapper jom=new JSONObjectMapper(envName, excelFileName);
					HashMap<String, String> result = jom.updateEnvConfig();
					request.setAttribute("result", result);
					request.setAttribute("action", "update");
		            dispatcher = getServletContext().getRequestDispatcher("/mapperresult.jsp");
				}
			} else if(action.equalsIgnoreCase("detailCompare")) {
				String apiName = request.getParameter("apiName");
				JSONComparator jc = new JSONComparator(apiName);
				LinkedHashMap<String, LinkedHashMap<String,String>> dispMap = jc.detailCompare();
			    request.setAttribute("diffResult", dispMap);
	            dispatcher = getServletContext().getRequestDispatcher("/filecomparison.jsp");
			}else if(action.equalsIgnoreCase("quickCompare")) {
				String apiName = request.getParameter("apiName");
				JSONComparator jc = new JSONComparator(apiName);
				LinkedHashMap<String, LinkedHashMap<String,String>> dispMap = jc.quickCompare();
			    request.setAttribute("diffResult", dispMap);
	            dispatcher = getServletContext().getRequestDispatcher("/filecomparison.jsp");
			}
            dispatcher.forward(request, response);
		}catch (Exception e) {
			request.setAttribute("error", "error");
			dispatcher = getServletContext().getRequestDispatcher("/jsonmapper.jsp");
            dispatcher.forward(request, response);
		}
	}

}
