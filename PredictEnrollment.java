import java.io.*;

import java.util.*;
import java.util.Set;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import java.text.DecimalFormat;

import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.ResultSet;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.PosixFileAttributes;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import static java.nio.file.attribute.PosixFilePermission.*;

import com.opencsv.CSVParser;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

import weka.core.Instance;
import weka.core.Instances;
import weka.core.Attribute;
import weka.core.converters.CSVLoader;
import weka.core.converters.CSVSaver;
import weka.core.converters.ConverterUtils.DataSource;

import weka.classifiers.Classifier;

import weka.classifiers.functions.GaussianProcesses;
import weka.classifiers.functions.LinearRegression;
import weka.classifiers.functions.MultilayerPerceptron;
import weka.classifiers.functions.SMOreg;

import weka.classifiers.evaluation.NumericPrediction;

import weka.classifiers.timeseries.AbstractForecaster;
import weka.classifiers.timeseries.core.TSLagMaker;
import weka.classifiers.timeseries.eval.TSEvaluation;
import weka.classifiers.timeseries.WekaForecaster;

import org.rosuda.JRI.REXP;
import org.rosuda.JRI.Rengine;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.CellRangeAddress;

public class PredictEnrollment {

  public static void main(String[] args) {

       /************
        * System:  MySQL
	* HOST:    mysql-prod.ptg.csun.edu
	* ReadOnly: 'r_prediction', 'ToViYiBG'
        ***********/
  
    try {
      Connection mysqlConnection = DriverManager.getConnection("jdbc:mysql://mysql-prod.ptg.csun.edu:50075/prediction","r_prediction","ToViYiBG"); 
      Statement stmt = mysqlConnection.createStatement();
      stmt.executeQuery("USE prediction;"); 
      stmt.executeQuery("SET group_concat_max_len=990000;"); 
      stmt.executeQuery("SET @sql = NULL;"); 
      /* Get enrollment data for all undergraduate computer science courses for all academic terms and years that have begin dates before today's date */
      /* Do Not add any whitespace to the string arguments to executeQuery - it will result in errors */
      stmt.executeQuery("SELECT GROUP_CONCAT(DISTINCT CONCAT('COUNT(DISTINCT(CASE WHEN sct.SUBJECT=''', sct.SUBJECT, ''' AND sct.CATALOG_NBR=''', sct.CATALOG_NBR, ''' THEN sse.EMPLID END)) AS `', sct.SUBJECT, sct.CATALOG_NBR, '_Enrollment_Total`')) INTO @sql FROM SA_CLASS_TBL sct WHERE sct.DEPT_NAME='Computer Science' AND sct.CLASS_STS='A' AND CAST(LEFT(sct.CATALOG_NBR, 1)AS UNSIGNED)<5"); 
      stmt.executeQuery("SET @sql = CONCAT('SELECT stt.BEGIN_DATE AS StartDate, ', @sql, ' FROM SA_CLASS_TBL sct LEFT JOIN SA_STDNT_ENRLS sse ON sct.STRM=sse.STRM AND sct.CLASS_NBR=sse.CLASS_NBR LEFT JOIN SA_TERM_TBL stt ON stt.STRM=sct.STRM WHERE stt.BEGIN_DATE<NOW() AND sct.DEPT_NAME=''Computer Science'' AND sct.CLASS_STS=''A'' AND CAST(LEFT(sct.CATALOG_NBR, 1) AS UNSIGNED)<5 GROUP BY stt.BEGIN_DATE, stt.TERM')"); 
      stmt.executeQuery("PREPARE stmt FROM @sql;"); 
      ResultSet queryResult = stmt.executeQuery("EXECUTE stmt;"); 

      /* Write the mySQL query results including the header to a csv file */
      /* The csv file will be the data input to Weka and R for enrollment forecasting */
      CSVWriter writer = new CSVWriter(new FileWriter(new File("/tmp/CoursePredictionSQLQueryOutput.csv"))); 
      writer.writeAll(queryResult,true);
      writer.flush();
      writer.close();

      stmt.executeQuery("DEALLOCATE PREPARE stmt;"); 
      stmt.close();

      String pathToData = "/tmp/CoursePredictionSQLQueryOutput.csv"; 
      String pathToCleanedData = "/tmp/CoursePredictionCleaned.csv"; 
      String pathToOutFile = "CourseEnrollmentPredictions.xlsx"; 

      /* Get all of the courses to predict enrollment for from the header of the CoursePredictionSQLQueryOutput.csv starting at index 1 (indices start at 0) */
      CSVReader reader = new CSVReader(new FileReader(pathToData));
      String[] rowHeader = reader.readNext(); 
      reader.close();

      String allcourses[] = new String[rowHeader.length-1]; 
      for(int i=1; i<rowHeader.length; i++){
         allcourses[i-1] = rowHeader[i];
      }

      /* Load the csv enrollment data without any overlay fields */
      /* First column is the date, all subsequent columns are numeric total enrollment, memory buffer size is increased */
      CSVLoader loader = new CSVLoader();
      loader.setSource(new File(pathToData));
      loader.setDateAttributes("1");

      /* JDBC changes the timestamp format to this format */
      loader.setDateFormat("dd-MMM-yyyy HH:mm:ss");
      loader.setBufferSize(100000);
      Instances enrollment = loader.getDataSet();

      /* Do not predict when a course has only enrollment of 0 by removing columns that only have 0 for their value */
      /* SMOreg will not predict enrollment for a course when a course has the same enrollment number for all of the data */
      /* Since it is unlikely for 5+ years of data for a course to have the same non-zero enrollment we only look for the zero value */
      /* Since 0 would indicate the course is rarely offered and so is a likely occurrence */ 
      /* Note that the number of instances (aka rows in the csv) is correct because the header isn't counted but the */
      /* number of attributes (aka columns in the csv) is 1 greater because it is counting the */
      /* first column which is the date which we want to ignore */ 
      int[] nonZeroEnrollment = new int[allcourses.length];
      for(int i=0; i<enrollment.numInstances(); i++){
         Instance semester = enrollment.instance(i);
         for(int j=1; j<semester.numAttributes(); j++){
           Attribute course = semester.attribute(j);
           if(semester.value(course)>0){
              nonZeroEnrollment[j-1] = 1;
           }
         }
      }

      /* Needed since when an attribute is removed, all subsequent attribute's indices decrease by 1 */
      int decrement = 0;
      for(int n=0; n<nonZeroEnrollment.length; n++){
         if(nonZeroEnrollment[n] == 0){
            /* Remove the attribute for each instance (ie remove the course values for each semester) */
            /* Index is +1 because the first attribute is the date */
            enrollment.deleteAttributeAt(n-decrement+1);
            decrement++;
         }
      }

      /* Write the cleaned data to a new csv file */
      /* Use Weka's util since the Instances obj is in .arff format and we want csv */
      CSVSaver saver = new CSVSaver(); 
      saver.setInstances(enrollment); 
      saver.setFile(new File(pathToCleanedData)); 
      saver.writeBatch(); 

      /* Update the header and courses array to reflect the clean data */
      reader = new CSVReader(new FileReader(pathToCleanedData));
      rowHeader = reader.readNext(); 
      reader.close();

      String courses[] = new String[rowHeader.length-1]; 
      for(int i=1; i<rowHeader.length; i++){
         courses[i-1] = rowHeader[i];
      }

      /* Predict the enrollment for next two upcoming semesters where semesters are Fall, Spring, Summer */ 
      /* since no undergraduate computer science courses are offered over the Winter semester */
      int numSemestersToForecast = 3;

      /* Do not hold out any data when making the classifier for testing */ 
      /* The more data used to build the classifier = the more accurate the classifier will be */
      float percentHoldOut = 0.0f;

      /* Round predictions to 2 decimal places */
      DecimalFormat df = new DecimalFormat("#0.##");

      /* The row to start inserting prediction data at */
      /* Begin at 1 since the header is the first row and indices start at 0 */
      int rowP = 1;

      /* The row to start inserting accuracy data at */
      /* Begin at 2 since headers are the first two rows and indices start at 0 */
      int rowA = 2;

      /* The number of classifiers to process using R = Arima, ETS, and RWF models */
      int rModelCnt = 3;

      /* Classifiers to feed to Weka */
      Classifier[] classifiers = {new GaussianProcesses(), new LinearRegression(), new MultilayerPerceptron(), new SMOreg()}; 

      String[] headerA = {"Courses", "GaussianProcesses", "LinearRegression", "MultilayerPerceptron", "SMOreg", "Arima", "ETS", "RWF"};

      String[] headerP = new String[(headerA.length-1)*numSemestersToForecast+1];
      headerP[0] = "Courses";
      int step = 0;
      for(int i=1; i<headerA.length; i++){
         for(int j=0; j<numSemestersToForecast; j++){
            headerP[j+1+step] = headerA[i] + "-" + (j+1) + "_Sem_Ahead"; 
         }
         step += numSemestersToForecast;
      }


      /* There must be a minimum of two accuracy methods due to cell merges in formatting */
      /* Due to having to parse formatted output to get the accuracy values more code than the strings here would need to be changed */
      String[] accuracy = {"MAE", "RMSE"};

      /* Create Excel Workbook (.xlsx) to write results to different spreadsheets */
      XSSFWorkbook workbook = new XSSFWorkbook();
      XSSFSheet sheetP = workbook.createSheet("CourseEnrollmentPredictions");
      XSSFSheet sheetA = workbook.createSheet("ForecastAccuracy");

      File outfile = new File(pathToOutFile);

      FileOutputStream fileout = new FileOutputStream(outfile);

      /* Set permissions on the xlsx workbook */
      outfile.setReadable(true, false);
      outfile.setWritable(true, false);
      outfile.setExecutable(true, false);

      /* Set permissions on the xlsx workbook using POSIX (won't work on Windows) */
      Set<PosixFilePermission> perms = new HashSet<>();
      perms.add(OWNER_READ);
      perms.add(OWNER_WRITE);
      perms.add(OWNER_EXECUTE);
      perms.add(GROUP_READ);
      perms.add(GROUP_WRITE);
      perms.add(GROUP_EXECUTE);
      perms.add(OTHERS_READ);
      perms.add(OTHERS_WRITE);
      perms.add(OTHERS_EXECUTE);
      Files.setPosixFilePermissions(Paths.get(pathToOutFile), perms);

      workbook.write(fileout);
      fileout.flush();
      fileout.close();

      /* BEGIN R FORECAST CODE */
      Rengine re = new Rengine (new String [] {"--vanilla"}, false, null);

       if (!re.waitForR())
       {
           System.out.println ("Unable to load R");
           return;
       }
       else {
           System.out.println ("Connected to R\n");

           /* Generate a log file for any R errors that are thrown */
           //re.eval("log<-file('R_Logfile.txt');");
           re.eval("sink(log, append=TRUE);");
           re.eval("sink(log, append=TRUE, type='message');");

           re.eval("library(zoo);");
           re.eval("library(timeDate);");
           re.eval("library(rJava);");
           re.eval("library(openxlsx);");
           re.eval("library(forecast);");
           re.eval("library(reshape2);");

           re.eval("data<-read.csv(file='/tmp/CoursePredictionCleaned.csv');");

           /* Convert the data into a time series and predict on the non-date, course enrollment columns */
           /* The date is the first column in the data so remove it to predict on the subsequent columns */
           /* The frequency corresponds to the seasonality of the data, where 3 represents the 3 semesters courses are offered = fall, spring, summer */
           re.eval("datats<-ts(data[,-1],frequency=3);");

           /* Number of columns representing courses to predict enrollment for */
           re.eval("ncc<-ncol(datats);");

           /* Number of columns for accuracy */
           /* R by default outputs 7 (ME, RMSE, MAE, MPE, MAPE, MASE, ACF1) */
           re.eval("nca<-7;");

           /* The accuracy measurements to write to the Excel spreadsheet */
           re.eval("accmeasures<-c('MAE','RMSE');");
           re.eval("laccm<-length(accmeasures);");

           /* Number of columns for forecast items output by default by R */ 
           /* PointForecast, Lo80, Hi80, Lo95, Hi95 (PointForecast = $mean) */
           re.eval("ncf<-5;");

           /* Number of predictions to make after the end of the existing data */
           /* Each prediction is for a semester, where semesters are Fall, Spring, and Summer */
           /* Due to Excel formatting changing this will require changing Excel formatting code */
           re.eval("h<-3;");

           /* Number of models (aka Classifiers) used = auto.arima, ets, rwf */
           re.eval("nmu<-3;");

           /* Store the PointForecasts ($mean) in matrices */
           re.eval("fcast_arima<-matrix(NA,nrow=h,ncol=ncc);");
           re.eval("fcast_ets<-matrix(NA,nrow=h,ncol=ncc);");
           re.eval("fcast_rwf<-matrix(NA,nrow=h,ncol=ncc);");

           /* accuracy() returns in-sample ME, RMSE, MAE, MPE, MAPE, MASE, ACF1 measurements */
           /* Will need to pick off the columns we want to put in the Excel spreadsheet (MAE and RMSE) */
           re.eval("facc_arima<-list();");
           re.eval("facc_ets<-list();");
           re.eval("facc_rwf<-list();");
           re.eval("for(i in 1:ncc){facc_arima[[i]]<-matrix(NA,nrow=1,ncol=nca)};");
           re.eval("for(i in 1:ncc){facc_ets[[i]]<-matrix(NA,nrow=1,ncol=nca)};");
           re.eval("for(i in 1:ncc){facc_rwf[[i]]<-matrix(NA,nrow=1,ncol=nca)};");

           /* Store the full forecast (ffc) as a list of 1x5 matrices = (in_sample_accuracy_results)x(forecast items: PointForecast, Lo80, Hi80, Lo95, Hi95) */
           /* It can be expanded to 2x5 when there is a test data passed to accuracy, then the 2nd row would be test accuracy results */
           re.eval("ffc_arima<-list();");
           re.eval("ffc_ets<-list();");
           re.eval("ffc_rwf<-list();");
           re.eval("for(i in 1:ncc){ffc_arima[[i]]<-matrix(NA,nrow=1,ncol=ncf)};");
           re.eval("for(i in 1:ncc){ffc_ets[[i]]<-matrix(NA,nrow=1,ncol=ncf)};");
           re.eval("for(i in 1:ncc){ffc_rwf[[i]]<-matrix(NA,nrow=1,ncol=ncf)};");
 
           /* forecast() returns PointForecast, Lo80, Hi80, Lo95, Hi95 (where PointForecast = $mean) */
           re.eval("for(i in 1:ncc){ffc_arima[[i]]<-forecast(auto.arima(datats[,i],approximation=FALSE,trace=FALSE),h=h)};");
           re.eval("for(i in 1:ncc){ffc_ets[[i]]<-forecast(ets(datats[,i]),h=h)};");
           re.eval("for(i in 1:ncc){ffc_rwf[[i]]<-forecast(rwf(datats[,i],h=h,drift=TRUE),h=h)};");

           /* Extract the PointForecast ($mean) from the full forecast (ffc) and round to 2 decimal places */
           re.eval("for(i in 1:ncc){fcast_arima[,i]<-round(ffc_arima[[i]]$mean,2)};");
           re.eval("for(i in 1:ncc){fcast_ets[,i]<-round(ffc_ets[[i]]$mean,2)};");
           re.eval("for(i in 1:ncc){fcast_rwf[,i]<-round(ffc_rwf[[i]]$mean,2)};");

           /* Round the training set (aka in-sample) accuracy measurements to 2 decimal places */
           re.eval("for(i in 1:ncc){facc_arima[[i]]<-round(accuracy(ffc_arima[[i]]),2)};");
           re.eval("for(i in 1:ncc){facc_ets[[i]]<-round(accuracy(ffc_ets[[i]]),2)};");
           re.eval("for(i in 1:ncc){facc_rwf[[i]]<-round(accuracy(ffc_rwf[[i]]),2)};");
           
           /* Open the Excel Workbook to write to Predictions and Accuracy Measurements sheets */
           re.eval("file.exists('CourseEnrollmentPredictions.xlsx');");
           re.eval("outwb<-loadWorkbook('CourseEnrollmentPredictions.xlsx',xlsxFile=NULL);");

           /* Write the forecast prediction results to Predictions sheet */
           /* Output the prediction values */
           re.eval("writeData(outwb,sheet=1,t(fcast_arima),startRow=1,startCol=1,colNames=FALSE,rowNames=FALSE);");
           re.eval("writeData(outwb,sheet=1,t(fcast_ets),startRow=1,startCol=1+h,colNames=FALSE,rowNames=FALSE);");
           re.eval("writeData(outwb,sheet=1,t(fcast_rwf),startRow=1,startCol=1+(2*h),colNames=FALSE,rowNames=FALSE);");

           /* Add MAE data for each classifier to sheetAcc spreadsheet */
           /* Index to access RMSE from R output is 2, Index to access MAE from R output is 3, Indices start at 1 */
           re.eval("for(i in 1:ncc){writeData(outwb,sheet=2,data.frame(facc_arima[[i]][3],facc_arima[[i]][2]),startRow=i,startCol=1,colNames=FALSE,rowNames=FALSE)};");
           re.eval("for(i in 1:ncc){writeData(outwb,sheet=2,data.frame(facc_ets[[i]][3],facc_ets[[i]][2]),startRow=i,startCol=1+laccm,colNames=FALSE,rowNames=FALSE)};");
           re.eval("for(i in 1:ncc){writeData(outwb,sheet=2,data.frame(facc_rwf[[i]][3],facc_rwf[[i]][2]),startRow=i,startCol=1+(2*laccm),colNames=FALSE,rowNames=FALSE)};");

           /* Save the Excel Workbook */
           re.eval("saveWorkbook(outwb,'CourseEnrollmentPredictions.xlsx',overwrite=TRUE);");
       }
       re.end();
       /* END OF R CODE SECTION */

      /* Read R predictions and accuracies from Excel spreadsheet before adding headers or Weka data */
      /* R data must be appended to Weka data before writing add data to spreadsheet since writing overwrites all existing data */
      FileInputStream inputStream = new FileInputStream(new File(pathToOutFile));
      XSSFWorkbook readWorkbook = new XSSFWorkbook(inputStream);

      /* Get prediction data from first sheet */
      XSSFSheet predSheet = readWorkbook.getSheetAt(0);

      Row[] rPredRows = new Row [courses.length];
      int rowIdx = 0;

      /* Go through prediction sheet row by row and store the R forecast data */
      Iterator<Row> predRowIterator = predSheet.iterator();
      while (predRowIterator.hasNext()) {
         rPredRows[rowIdx] = predRowIterator.next();
         rowIdx++;
      }

      /* Get accuracy data from second sheet */
      XSSFSheet accSheet = readWorkbook.getSheetAt(1);

      Row[] rAccRows = new Row [courses.length];
      rowIdx = 0;

      /* Go through accuracy sheet row by row and store the R forecast data */
      Iterator<Row> accRowIterator = accSheet.iterator();
      while (accRowIterator.hasNext()) {
         rAccRows[rowIdx] = accRowIterator.next();
         rowIdx++;
      }

      /* Close the connection to the Excel workbook since done reading */
      inputStream.close();

      /* Create Workbook sheet header style */
      XSSFCellStyle headerStyle = workbook.createCellStyle();
      headerStyle.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
      headerStyle.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
      headerStyle.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
      headerStyle.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);

      /* Create Workbook sheet header font */
      XSSFFont headerFont = workbook.createFont();
      headerFont.setFontHeightInPoints((short) 12);
      headerFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
      headerStyle.setFont(headerFont); 

      /* Map to store the prediction and accuracy data to output to the Excel spreadsheets */
      Map<String, Object[]> dataP = new TreeMap<String, Object[]>();
      Map<String, Object[]> dataA = new TreeMap<String, Object[]>();

      /* Object[Rows][Columns] where */
      /* Rows = the courses for which to predict the total enrollment for numSemestersToForecast steps ahead in time */ 
      /* Columns = the predictions for each step ahead in time for each classifier in classifiers[] */
      Object[][] predictions = new Object[courses.length][classifiers.length*numSemestersToForecast];
      Object[][] accuracies = new Object[courses.length][classifiers.length*accuracy.length];

      /* Add the CourseEnrollmentPredictions sheet column headers */
      XSSFRow hrowP1 = sheetP.createRow(0);
      for(int col=0; col<headerP.length; col++){
         XSSFCell hcell = hrowP1.createCell(col);
         hcell.setCellValue(headerP[col]);
         hcell.setCellStyle(headerStyle);
      }

      /* Add the ForecastAccuracy sheet column headers */
      XSSFRow hrowA1 = sheetA.createRow(0);
      XSSFRow hrowA2 = sheetA.createRow(1);
      for(int row=0; row<rowA; row++){
         /* Add 1 in the comparison check due to the first column header being 'Courses' and not a classifier */
         for(int col=0; col<(1+((headerA.length-1)*accuracy.length)); col++){
            /* 1st column is 'Courses' and not a classifier so it is not merged */
            if(col==0){
               if(row==0){
                  XSSFCell hcell = hrowA1.createCell(col);
                  hcell.setCellValue(headerA[col]);
                  hcell.setCellStyle(headerStyle);
               }
               else{ /* Do nothing because row=1, col=0 is empty */ }
            }
            else{
               /* First header classifiers start at column index 1 */
               /* Each classifier name is in a merged cell of size accuracy.length */
               if(row==0){
                  if(col==1 || (col%accuracy.length==1)){
                     CellRangeAddress range = new CellRangeAddress(0,0,col,col+(accuracy.length-1)); 
                     sheetA.addMergedRegion(range);
                     XSSFCell hcell = hrowA1.createCell(col);
                     int index = (col/accuracy.length) + 1;
                     hcell.setCellValue(headerA[index]);
                     hcell.setCellStyle(headerStyle);
                     RegionUtil.setBorderRight(XSSFCellStyle.BORDER_MEDIUM, range, sheetA, workbook);
                     RegionUtil.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM, range, sheetA, workbook);
                     RegionUtil.setBorderTop(XSSFCellStyle.BORDER_MEDIUM, range, sheetA, workbook);
                     RegionUtil.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM, range, sheetA, workbook);
                  }
               }
               /* Second header of accuracy measurements starts at column index 1 */
               else {
                  XSSFCell hcell = hrowA2.createCell(col);
                  int index = (col+1)%accuracy.length;
                  hcell.setCellValue(accuracy[index]);
                  hcell.setCellStyle(headerStyle);
               }
            }
         }
      }
     
      /* Forecaster to get predicted values */
      WekaForecaster forecaster = new WekaForecaster();

      int minLag = 1;
      int maxLag = 3;
    
      int j = 0;
      int s = 0;
 
      /* Get predictions and accuracy for each classifier and generate an Excel spreadsheet with the results */
      for(int i=0; i<classifiers.length; i++){

         /* Set the targets we want to forecast */
         forecaster.setFieldsToForecast(String.join(",", courses));

         forecaster.setBaseForecaster(classifiers[i]);
         forecaster.setConfidenceLevel(0.95);

         /* StartDate corresponds to SA_TERM_TBL.BEGIN_DATE */
         forecaster.getTSLagMaker().setTimeStampField("StartDate");
         forecaster.getTSLagMaker().setMinLag(minLag);
         forecaster.getTSLagMaker().setMaxLag(maxLag);

         /* Add month and quarter of the year indicator field */
         forecaster.getTSLagMaker().setAddMonthOfYear(true);
         forecaster.getTSLagMaker().setAddQuarterOfYear(true);

         /* Build the model */
         forecaster.buildForecaster(enrollment, System.out);

         /* Run the model with recent historical data */
         /* There must be at least maxLag amount of data */
         forecaster.primeForecaster(enrollment);

         /* Forecast for numSemestersToForecast units beyond the end of the training data without overlay data */
         List<List<NumericPrediction>> forecast = forecaster.forecast(numSemestersToForecast, System.out);
      
         /* Output the predictions */
         /* Outer j loop is over the steps forward in time & is the Excel spreadsheet columns */
         /* Inner k loop is over the target courses which are the different courses for which to predict enrollment & is the Excel spreadsheet rows */
         for(j=j; j<((numSemestersToForecast*i)+numSemestersToForecast); j++) {
            List<NumericPrediction> predsAtStep = forecast.get(j%numSemestersToForecast);

            for (int k=0; k<courses.length; k++) {
               NumericPrediction predForTarget = predsAtStep.get(k);

               /* Fill predictions column by column, storing with 2 decimal points of accuracy as Float objects */
               predictions[k][j] = new Float(df.format(predForTarget.predicted())); 
            }
         }

         /* Compute the evaluation metrics */
         TSEvaluation eval = new TSEvaluation(enrollment, percentHoldOut);
         eval.setHorizon(numSemestersToForecast);

         /* Get the Time Series Evaluation Summaries */
         eval.evaluateForecaster(forecaster, System.out);
         String accuracyMeasures = eval.toSummaryString();

         Pattern patternMAE = Pattern.compile("Mean absolute error.+");
         Pattern patternRMSE = Pattern.compile("Root mean squared error.+");
  
         Matcher matcherMAE = patternMAE.matcher(accuracyMeasures); 
         Matcher matcherRMSE = patternRMSE.matcher(accuracyMeasures);

         /* To keep track of which find (aka course) the accuracy measure is for */
         int k = 0;

         /* Each match in String accuracyMeasures is for a forecasted field (a course) for the current classifier */
         while(matcherMAE.find()){
            /* Match will look like: "Mean absolute error   330.1107   320.5035" */
            String groupMAE = matcherMAE.group();

            /* Extract the MAE numbers for each forecasted step (total of numSemestersToForecast steps) from the string */
            Pattern extractMAE = Pattern.compile("\\d+(\\.\\d+)?");
            Matcher matchMAE = extractMAE.matcher(groupMAE);

            float maeTotal = 0.0f;
            while(matchMAE.find()){
               String mae = matchMAE.group(); 
               maeTotal += Float.parseFloat(mae);
            }
       
            /* Store the average MAE across all forecasted steps with 2 decimal points of accuracy as Float objects */
            accuracies[k][s] = new Float(df.format(maeTotal/numSemestersToForecast));
            k++;
         }
        
         /* Reset for the next find loop */
         k = 0;
 
         while(matcherRMSE.find()){
            /* Match will look like: "Root mean squared error   415.7785   399.6755" */
            String groupRMSE = matcherRMSE.group();

            /* Extract the RMSE numbers for each forecasted step (total of numSemestersToForecast steps) from the string */
            Pattern extractRMSE = Pattern.compile("\\d+(\\.\\d+)?");
            Matcher matchRMSE = extractRMSE.matcher(groupRMSE);

            float rmseTotal = 0.0f;
            while(matchRMSE.find()){
               String rmse = matchRMSE.group(); 
               rmseTotal += Float.parseFloat(rmse);
            }

            /* Store the RMSE in the column after the MAE column with 2 decimal points of accuracy as Float objects */
            accuracies[k][s+1] = new Float(df.format(rmseTotal/numSemestersToForecast));
            k++;
         }
       
         /* Increment the column index by the number of accuracy measures - MAE & RMSE - so by 2 */ 
         s += accuracy.length; 

         /* Reset the forecaster */
         forecaster.reset();

      } /* END for loop that calculates predictions and accuracy of the classifiers */
      

      /* Write predictions row by row, where each row is predictions for an undergraduate COMP or CIT course */
      /* The index into the row[] can be thought of as the column in that row */
      for(int m=0; m<courses.length; m++) {
         /* Prepend the Course column data to each prediction row */
         /* Use classifiers.length+rModelCnt to get the total count of all models (Weka and R) */
         Object[] row = new Object [1+((classifiers.length+rModelCnt)*numSemestersToForecast)];

         /* row[0] is a String containing the course name, for example "COMP 100" */
         row[0] = courses[m];

         /* row[1] to row[classfiers.length*numSemestersToForecast] contain the Weka model predictions */
         for(int n=0; n<(classifiers.length*numSemestersToForecast); n++){
            row[n+1] = predictions[m][n];
         }

         /* row[classifiers.length*numSemestersToForecast + 1] to row[1+((classifiers.length+rModelCnt)*numSemestersToForecast)] */
         /* contain the R forecast model predictions */
         Iterator<Cell> predCellIterator = rPredRows[m].cellIterator();
         rowIdx = (classifiers.length*numSemestersToForecast) + 1;
         while (predCellIterator.hasNext()) {
            Cell rPredCell = predCellIterator.next();
            row[rowIdx++] = new Float(rPredCell.getNumericCellValue());
         }

         /* Insert the row after the header in the Excel spreadsheet */
         dataP.put(Integer.toString(rowP+m), row);
      }

      /* Write accuracies row by row, where each row is the average accuracy measurements of MAE and RMSE */
      /* for all predicted steps for each undergraduate COMP or CIT course */
      /* The index into the row[] can be thought of as the column in that row */
      for(int m=0; m<courses.length; m++) {
         /* Prepend the Course column data to each accuracy row */
         /* Use classifiers.length+rModelCnt to get the total count of all models (Weka and R) */
         Object[] row = new Object [1+((classifiers.length+rModelCnt)*accuracy.length)];
         
         /* row[0] is a String containing the course name, for example "COMP 100" */
         row[0] = courses[m];

         /* row[1] to row[classifiers.length*accuracy.length] contain accuracy data for Weka models */
         for(int n=0; n<(classifiers.length*accuracy.length); n++){
            row[n+1] = accuracies[m][n];
         }

         /* row[classifiers.length*accuracy.length] to row[1+((classifiers.length+rModelCnt)*accuracy.length)] */
         /* contain accuracy data for R forecast models */
         Iterator<Cell> accCellIterator = rAccRows[m].cellIterator();
         rowIdx = (classifiers.length*accuracy.length) + 1;
         while (accCellIterator.hasNext()) {
            Cell rAccCell = accCellIterator.next();
            row[rowIdx++] = new Float(rAccCell.getNumericCellValue());
         }

         /* Insert the row after the header in the Excel spreadsheet */
         dataA.put(Integer.toString(rowA+m), row);
      }

      /* Write the prediction data to the first Excel spreadsheet CourseEnrollmentPredictions */
      /* Data starts at Row 1, Column 0 (indices start at 0) */
      Set<String> keysetP = dataP.keySet();
      int rownumP = rowP;
      for(String key : keysetP){
         XSSFRow row = sheetP.createRow(rownumP++);
         Object[] objArr = dataP.get(key);
   
         /* Add cells to the row column by column */
         int cellnumP = 0;
         for(Object obj : objArr){
            XSSFCell cell = row.createCell(cellnumP++);
            /* The predictions are of type Float */
            if(obj instanceof Float){
               cell.setCellValue((Float)obj);
            }
            /* The course names are of type String */
            else {
               /* Treat cell contents as type String by default */
               cell.setCellValue((String)obj);
            }
         }
      }

      /* Write the accuracy data to the second Excel spreadsheet named: ForecastAccuracy.xlsx */
      /* Data starts at Row 2, Column 1 (indices start at 0) */
      Set<String> keysetA = dataA.keySet();
      int rownumA = rowA;
      for(String key : keysetA){
         XSSFRow row = sheetA.createRow(rownumA++);
         Object[] objArr = dataA.get(key);
         int cellnumA = 0;
         for(Object obj : objArr){
            XSSFCell cell = row.createCell(cellnumA++);
            /* The accuracies are of type Float */
            if(obj instanceof Float){
               cell.setCellValue((Float)obj);
            }
            /* The course names are of type String */
            else {
               /* Treat cell contents as type String by default */
               cell.setCellValue((String)obj);
            }
         }
      }

      /* Write the Headers, Weka, and R output to the xlsx workbook */
      /* The cells that R data will be written to are initialized to an empty string as required by openxlsx's loadWorkbook */ 
      outfile = new File(pathToOutFile);
      fileout = new FileOutputStream(outfile);
      workbook.write(fileout);
      fileout.flush();
      fileout.close();
    }
    catch (Exception ex) {
      ex.printStackTrace();
    }
  }
}
