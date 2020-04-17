package com.mvc.poi.service.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.mvc.poi.dao.ExcelDao;
import com.mvc.poi.service.ExcelService;
import com.mvc.poi.vo.EmpVO;

@Service("excelService")
public class ExcelServiceImpl implements ExcelService{

	@Resource(name="excelDao")
	private ExcelDao excelDao;
	
	@Override
	public List<EmpVO> getEmpList() {
		// TODO Auto-generated method stub
		 return excelDao.getEmpList();
	}
}
