# easyexcelutilexample
easyexcelutil使用方法（这里只是把阿里的easyexcel做个个简化封装）
项目引入
````
<dependency>
    <groupId>com.github.zywaiting</groupId>
    <artifactId>easyexcelutil</artifactId>
    <version>0.0.1</version>
</dependency>
````
#test是测试没什么用,实体类要参考
（阿里的）
EasyExcel 的 github 地址： 
https://github.com/alibaba/easyexcel


##object 实体类映射，继承 BaseRowModel 类
###读取 Excel(多个 sheet)

####MultipartFile  
readExcel(MultipartFile excel, BaseRowModel object)
####File
readExcel(File excel, BaseRowModel object)
####fileName 文件名字 InputStream 输入流
readExcel(String fileName, InputStream inputStream, BaseRowModel object)

##读取某个 sheet 的 Excel
###XLS 类型文件 sheet 序号为顺序，第一个 sheet 序号为 1
###XLSX 类型 sheet 序号顺序为倒序，即最后一个 sheet 序号为 1

readExcel(MultipartFile excel, BaseRowModel object, int sheetNo)

readExcel(File excel, BaseRowModel object, int sheetNo)

readExcel(String fileName, InputStream inputStream, int sheetNo)


##导出 Excel ：一个或多个 sheet，带表头
fileName  导出的文件名
writeExcelWithSheets(HttpServletResponse response, String fileName)

````
/**
     * 导出 Excel（一个 sheet）
     */
    @RequestMapping(value = "writeExcel", method = RequestMethod.GET)
    public void writeExcel(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName = "第一个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName, new ExportInfo());
    }

````

````
/**
     * 导出 Excel（多个 sheet）
     */
    @RequestMapping(value = "writeExcelWithSheets", method = RequestMethod.GET)
    public void writeExcelWithSheets(HttpServletResponse response) throws IOException {
        List<ExportInfo> list = getList();
        String fileName = "一个 Excel 文件";
        String sheetName1 = "第一个 sheet";
        String sheetName2 = "第二个 sheet";
        String sheetName3 = "第三个 sheet";

        EasyExcelUtil.writeExcelWithSheets(response, fileName)
                .write(list, sheetName1, new ExportInfo())
                .write(list, sheetName2, new ExportInfo())
                .write(list, sheetName3, new ExportInfo())
                .finish();
    }
````



