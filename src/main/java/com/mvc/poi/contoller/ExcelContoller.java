package com.mvc.poi.contoller;

import java.util.List;
import java.util.Map;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.mvc.poi.service.ExcelService;
import com.mvc.poi.vo.EmpVO;

@Controller
public class ExcelContoller {

	@Resource(name="excelService")
	private ExcelService excelService;
	
	@RequestMapping("/excelDownload")
    public String excelTransform(Map<String, Object> model, HttpServletResponse response) throws Exception{
        
        response.setHeader("Content-disposition", "attachment; filename= empTest.xlsx");

        List<EmpVO> empList = excelService.getEmpList(); 
        
        model.put("empList", empList); 
 
        return "excelView";
 
    }
}
