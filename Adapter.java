public static String ExportFile (modelExport param1, mdlExportParam param2){

    String sql = "";
    String functionName = Thread.currentThread().getStackTrace()[1].getMethodName();
    String urlLink = "";
    String user = param1.created_by;
    Workbook workbook = new XSSFWorkbook();
    List<mdlQueryExecute> _mdlQueryExecuteList = new ArrayList<mdlQueryExecute>();
    CachedRowSet jrs = null;
    List<modelExport> modelExportList = new ArrayList<modelExport>();
    FileOutputStream outPutStream = null;

    try {

        sql = "{}"; // ==> YOUR QUERY
        jrs = QueryAdapter.QueryExecuteWithDB(sql, _mdlQueryExecuteList, functionName, user, Globals.dbName);

        while (jrs.next()) {
            modelExport _modelExport = new modelExport();
            _modelExport.ID = jrs.getString("ID");
            _modelExport.FK = jrs.getString("FK");
            _modelExport.Name = jrs.getString("Name");
            modelExportList.add(_modelExport);
        }

        //== Create Workbook==\\

        String[] columns = {"ID", "FK", "Name"};


        Sheet sheet = workbook.createSheet("File-Export");

        // Styling your font on header at excel file


        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 11);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        for(int i = 0; i < columns.length; i++ ){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }


        int rowNum = 1;
        for(modelExport _modelExport : modelExportList){
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(_modelExport.ID);
            row.createCell(1).setCellValue(_modelExport.FK);
            row.createCell(2).setCellValue(_modelExport.Name);
        }

        for (int i = 0; i < columns.length; i++){
            sheet.autoSizeColumn(i);
        }


        try {
            Context context = (Context) new InitialContext().lookup("java:comp/env");
            String paramFilePath = (String) context.lookup("path_web_server_for_upload_report");
            String paramFileUrl = (String) context.lookup("url_web_server_for_upload_report");

            String fileName = "Export_" + DateHelper.GetDateTimeNowCustomFormat("yyyyMMddHHmmss");
            String url = new String(paramFileUrl+ param2.app_name  +"/Export/" + fileName + ".xlsx");
            File createFile = new File(paramFilePath+ param2.app_name + "/Export");
            createFile.mkdirs();

            File myFile = new File(createFile.getAbsolutePath() + "/" + fileName + ".xlsx");

            outPutStream = new FileOutputStream(myFile);
            urlLink = url;
            workbook.write(outPutStream);
            outPutStream.close();
            workbook.close();


        }catch (Exception ex){
            ex.printStackTrace();
        }


    }catch (Exception ex){
        core.LogAdapter.InsertLogExc(ex.toString(), "Export", sql, user);
    }

    return urlLink;
}