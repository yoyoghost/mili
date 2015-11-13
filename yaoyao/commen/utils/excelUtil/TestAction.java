/**
	 * 不完全测试类
	 * 写在contrller层 
	 * web请求
	 */
public void ee(HttpServletResponse response){
		List<TestBean> list = ...;
		ExportExcelUtil excelUtil = new ExportExcelUtil();
		String[] excelTitle = excelUtil.getExcelTitle(TestBean.class);
		Map<Integer, List<Object>> result = excelUtil.getExcelContent(TestBean.class, list);
		excelUtil.initialize("", "", excelTitle);
		excelUtil.setContent(result);
		excelUtil.write(response,new Date().getTime()+".xlsx");
		excelUtil.dispose();
}
