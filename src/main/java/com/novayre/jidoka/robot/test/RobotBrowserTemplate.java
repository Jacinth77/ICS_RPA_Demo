package com.novayre.jidoka.robot.test;

import com.novayre.jidoka.client.api.*;
import com.novayre.jidoka.client.api.exceptions.JidokaFatalException;
import com.novayre.jidoka.client.api.exceptions.JidokaItemException;
import com.novayre.jidoka.client.api.exceptions.JidokaQueueException;
import com.novayre.jidoka.client.api.queue.*;
import com.novayre.jidoka.client.lowcode.IRobotVariable;
import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IJidokaDataProvider;
import com.novayre.jidoka.data.provider.api.IJidokaExcelDataProvider;
import com.novayre.jidoka.windows.api.IWindows;
import org.apache.commons.lang.StringUtils;

import com.novayre.jidoka.browser.api.EBrowsers;
import com.novayre.jidoka.browser.api.IWebBrowserSupport;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.multios.IClient;
import org.apache.commons.lang.exception.ExceptionUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;


/**
 * Browser robot template. 
 */
@Robot
public class RobotBrowserTemplate implements IRobot {

	/**
	 * URL to navigate to.
	 */
	private static final String HOME_URL = "https://www.appian.com";
	
	/** The JidokaServer instance. */
	private IJidokaServer<?> server;
	
	/** The IClient module. */
	private IClient client;
	
	/** WebBrowser module */
	private IWebBrowserSupport browser;

	/** Browser type parameter **/
	private String browserType = null;

	private QueueCommons queueCommons;

	private ExcelDSRow excelDSRow;

	private IJidokaExcelDataProvider<Excel_Input> dataProvider;

	private IExcel excel;

	private IQueueManager qmanager;

	private String selectedQueueID;

	private IQueue currentQueue;

	private Excel_Input excelinput;
	//private Excel_Input currentItem;

	private static final int FIRST_ROW = 0;

	private IQueueItem currentItemQueue;
	private int currentItemIndex=0;

	private ICS_WebApplication webApplication;

	private IWindows windows;

	private String returnType ="No";

	private String OutputFilepath;

	private Integer cellNumber = 2;

	private String Status = "Success";

	private OCR ocr;

	private String documentType;

	private String idNumber;

	private IKeyboard keyboard;

	/**
	 * Action "startUp".
	 * <p>
	 * This method is overrriden to initialize the Appian RPA modules instances.
	 */
	@Override
	public boolean startUp() throws Exception {
		
		server = (IJidokaServer< ? >) JidokaFactory.getServer();

		client = IClient.getInstance(this);

		browser = IWebBrowserSupport.getInstance(this, client);
		windows = IJidokaRobot.getInstance(this);

		return IRobot.super.startUp();

	}
	
	/**
	 * Action "start".
	 */
	public void start() throws Exception {
		server = (IJidokaServer< ? >) JidokaFactory.getServer();

		client = IClient.getInstance(this);
		ocr = new OCR();

		browser = IWebBrowserSupport.getInstance(this, client);
		qmanager = server.getQueueManager();
		keyboard=client.getKeyboard();
		queueCommons = new QueueCommons();
		webApplication =new ICS_WebApplication();
		excelDSRow = new ExcelDSRow();
		queueCommons.init(qmanager);
		dataProvider = IJidokaDataProvider.getInstance(this, IJidokaDataProvider.Provider.EXCEL);
		server.setNumberOfItems(1);
		excel = IExcel.getExcelInstance(this);
		server = (IJidokaServer< ? >) JidokaFactory.getServer();
		//excelinput= new Excel_Input();

		client = IClient.getInstance(this);

		browser = IWebBrowserSupport.getInstance(this, client);
		//return IRobot.super.startUp();
	}

    public void resetvariables(){
		server.info("Reset Variables");
		returnType ="No";
		webApplication.reset();
	}
	public String MaxCountReached(){
		if (returnType.contains("MaxCountReached")){
			Status = "Failed";
			return "Yes";
		}
		else if (webApplication.dict.isEmpty()){
			Status = "Employee Not Found";
			return "Yes";
		}
		else
		{
			return "No";
		}


	}
/**
	 * Navigate to Web Page
	 * 
	 * @throws Exception
	 */
	public void navigateToCustomerWeb() throws Exception  {
		//Set File Name for Customer Xpath
		String cXpathFileName = server.getEnvironmentVariables().get("customerXpathFileName").toString();
		server.info("InputID "+excelinput.getInput_ID());
		webApplication.PerformOperation(cXpathFileName,excelinput.getInput_ID());
		keyboard.altF(4);

	}
	public void navigateToGoogleWeb() throws Exception  {
		//Set File Name for Google Xpath
		String gXpathFileName = server.getEnvironmentVariables().get("GoogleXpathFileName").toString();
		server.info("InputID "+excelinput.getInput_ID());
		webApplication.PerformOperation(gXpathFileName,excelinput.getInput_ID());
		TimeUnit.SECONDS.sleep(5);
		keyboard.altF(4);


	}
	public String customerRetry() throws Exception
	{
		returnType = webApplication.RetryRequired();
		if (returnType.contains("Yes")) {
			return "Yes";
		} else  {
			return "No";
		}
	}
	public String googleRetry() throws Exception{
		returnType = webApplication.RetryRequired();
		if (returnType.contains("Yes")){
			return "Yes";
		}
		else{
			return "No";
		}
	}



	public void SelectQueue() throws Exception {
try {

	if (StringUtils.isNotBlank(qmanager.preselectedQueue())) {

		selectedQueueID = qmanager.preselectedQueue();
		server.info("Selected queue ID: " + selectedQueueID);
		addItemsToQueue();
		currentQueue = queueCommons.getQueueFromId(selectedQueueID);
	} else {

		String inputFilePath = server.getEnvironmentVariables().get("InputFilePath").toString().replace("*", "\\");
		selectedQueueID = queueCommons.createQueue(inputFilePath);
		server.info("Queue ID: " + selectedQueueID);
		addItemsToQueue();
		currentQueue = queueCommons.getQueueFromId(selectedQueueID);

	}
}
catch (Exception e){
	throw new JidokaQueueException("Unable to find / Create Queue" +e);
}

	}

	private void addItemsToQueue() throws Exception {


		String inputFilePath=server.getEnvironmentVariables().get("InputFilePath").toString().replace("*","\\");
		//Path inputFile = Paths.get(inputFilePath);


		dataProvider = IJidokaDataProvider.getInstance(this, IJidokaDataProvider.Provider.EXCEL);
		dataProvider.init(inputFilePath, "DataSource1", FIRST_ROW, new Excel_Input_RowMapper());

		server.info(inputFilePath);

		try {

			// Get the next row, each row is a item
			while (dataProvider.nextRow()) {

				server.info("inside while");

				CreateItemParameters itemParameters = new CreateItemParameters();
				Excel_Input  excelinput = dataProvider.getCurrentItem();



				// Set the item parameters
				itemParameters.setKey(excelinput.getInput_ID());
				itemParameters.setPriority(EPriority.NORMAL);
				itemParameters.setQueueId(selectedQueueID);
				itemParameters.setReference(String.valueOf(dataProvider.getCurrentItemNumber()));

				Map<String, String> functionalData = new HashMap<>();
				functionalData.put(Excel_Input_RowMapper.Input_ID,excelinput.getInput_ID());
				functionalData.put(Excel_Input_RowMapper.Status, excelinput.getStatus());

				itemParameters.setFunctionalData(functionalData);
				qmanager.createItem(itemParameters);

				server.debug(String.format("Added item to queue %s with id %s", itemParameters.getQueueId(),
						itemParameters.getKey()));

				server.setCurrentItem(currentItemIndex,"Success");
			}

		} catch (Exception e) {
			throw new JidokaItemException("Error While creating items");
		} finally {

			try {
				// Close the excel file
				dataProvider.close();
			} catch (IOException e) {
				throw new JidokaItemException("Add item to Queue- while close Excel"+e);
			}
		}
	}

	public String hasMoreItems() throws Exception {

		// retrieve the next item in the queue
		//QueueCommons queueCommons = new QueueCommons();
		currentItemQueue = queueCommons.getNextItem(currentQueue);

		if (currentItemQueue != null) {


            server.info("current item key"+currentItemQueue.functionalData().get("Input_ID"));
			excelinput= new Excel_Input();

			excelinput.setInput_ID(currentItemQueue.functionalData().get("Input_ID"));
			excelinput.setStatus(currentItemQueue.functionalData().get("Status"));

			server.info("Input Customer ID -"+excelinput.getInput_ID());

			/* set the stats for the current item
			Map<String, IRobotVariable> variables = server.getWorkflowVariables();
			IRobotVariable AN = variables.get("EmpNo");
			AN.setValue(excelinput.getInput_ID());*/

			server.setCurrentItem(currentItemIndex++, currentItemQueue.key());

			Status = "Success";
			webApplication.resetDictionary();

			return "Yes";
		}

		return "No";
	}

	public void releaseitems() throws Exception {

		if (webApplication.dict.isEmpty()){
			Status = "Employee Not Found";
		}

			updateInputExcel(Status);
			Map<String, String> funcData = currentItemQueue.functionalData();
			funcData.put(Excel_Input_RowMapper.Status,Status);
			ReleaseItemWithOptionalParameters rip = new ReleaseItemWithOptionalParameters();
			rip.functionalData(funcData);
			qmanager.releaseItem(rip);

	}

	public void encrptPDF() throws IOException {

		if (webApplication.dict.isEmpty() == false){

			File file = new File(OutputFilepath + ".pdf");
			PDDocument document = PDDocument.load(file);
			AccessPermission ap = new AccessPermission();
			String password = "ICS" + webApplication.dict.get("Emp");
			StandardProtectionPolicy spp = new StandardProtectionPolicy(password, password, ap);
			spp.setEncryptionKeyLength(128);
			spp.setPermissions(ap);
			document.protect(spp);
			server.info("Document encrypted");
			document.save(OutputFilepath + ".pdf");
			document.close();
		}
	}
	public void writeToExcel() throws Exception {

		server.info(webApplication.dict.isEmpty());
		if (webApplication.dict.isEmpty() == false){

			server.info("Inside WritetoExcel");

			String excelPath = Paths.get(server.getCurrentDir(), "FinalTemplate.xlsx").toString();
			//try (
			IExcel excelIns = IExcel.getExcelInstance(this);
			//{
				server.info("Excel Path" + excelPath);
				excelIns.init(excelPath);
				server.info("EmpValue" + webApplication.dict.get("Emp"));
				server.info("NameValue" + webApplication.dict.get("Name"));
				excelIns.setCellValueByName("I5", LocalDate.now().toString());
				excelIns.setCellValueByName("G10", webApplication.dict.get("Emp"));
				excelIns.setCellValueByName("G12", webApplication.dict.get("Name"));
				excelIns.setCellValueByName("G14", webApplication.dict.get("Designation"));
				excelIns.setCellValueByName("G16", webApplication.dict.get("Email id"));
				excelIns.setCellValueByName("G18", webApplication.dict.get("Mobile"));
				excelIns.setCellValueByName("G20", webApplication.dict.get("Project Code"));
				excelIns.setCellValueByName("G22", webApplication.dict.get("Practice Unit"));
				excelIns.setCellValueByName("G24", webApplication.dict.get("Current Location"));
				excelIns.setCellValueByName("G26", webApplication.dict.get("Current City"));
				excelIns.setCellValueByName("G28",documentType );

				server.info("IdNumber  "+idNumber);
				excelIns.setCellValueByName("G30",idNumber );
				server.info("End of Write");
				excelIns.close();

				Desktop.getDesktop().open(Paths.get(server.getCurrentDir(), "FinalTemplate.xlsx").toFile());
				TimeUnit.SECONDS.sleep(8);
				client.typeText(client.getKeyboardSequence().pressAlt().type("f").releaseAlt());
				TimeUnit.SECONDS.sleep(3);
				client.typeText(client.getKeyboardSequence().type("e"));
				TimeUnit.SECONDS.sleep(3);
				client.typeText(client.getKeyboardSequence().type("a"));
				TimeUnit.SECONDS.sleep(3);
				OutputFilepath = server.getEnvironmentVariables().get("OutPutFilePath").toString().replace("*", "\\") + "\\" + webApplication.dict.get("Emp") + " - " + webApplication.dict.get("Name");
				server.info("OutputFilepath  :" + OutputFilepath);
				client.typeText(client.getKeyboardSequence().type(OutputFilepath));
				windows.keyboard().enter();
				TimeUnit.SECONDS.sleep(4);
				keyboard.altF(4);

				//Runtime.getRuntime().exec("taskkill /F /IM EXCEL.exe");
				server.setCurrentItemResultToOK("Values Updated in Excel sheet");

			/*} catch (Exception e) {
				throw new JidokaItemException("Exception at  Write to Excel" );
			}*/
		}
	}

	public void closeQueue(){

	}
	public void GetData() throws Exception {


		String FilePath = server.getEnvironmentVariables().get("GoogleDocsPath").toString().replace("*","\\")+excelinput.getInput_ID()+".jpg";
		server.info("Filepath   "+FilePath);
		documentType=ocr.searchImage(FilePath);
		idNumber=ocr.ImgPreprocessing(FilePath);
	}

	private void updateInputExcel(String Status) throws Exception {

		String inputFilePath=server.getEnvironmentVariables().get("InputFilePath").toString().replace("*","\\");
		try (IExcel excelInsup = IExcel.getExcelInstance(this)) {
			server.info("Excel Path" + inputFilePath);
			excelInsup.init(inputFilePath);
			String cellName = "B" + cellNumber;

			excelInsup.setCellValueByName(cellName, Status);
			cellNumber = cellNumber + 1;
			server.info("End of Update");
			excelInsup.close();

		} catch (Exception e) {
			server.info(e);
		}
	}


			/**
             * @see com.novayre.jidoka.client.api.IRobot#cleanUp()
             */
	@Override
	public String[] cleanUp() throws Exception {
		
		browserCleanUp();
		return null;
	}

	/**
	 * Close the browser.
	 */
	private void browserCleanUp() {

		// If the browser was initialized, close it
		if (browser != null) {
			try {
				browser.close();
				browser = null;

			} catch (Exception e) { // NOPMD
			// Ignore exception
			}
		}

		try {
			
			if(browserType != null) {
				
				switch (EBrowsers.valueOf(browserType)) {

				case CHROME:
					client.killAllProcesses("chrome.exe", 1000);
					client.killAllProcesses("chromedriver.exe", 1000);
					break;

				case INTERNET_EXPLORER:
					client.killAllProcesses("IEDriverServer.exe", 1000);
					break;

				case FIREFOX:
					client.killAllProcesses("geckodriver.exe", 1000);
					break;

				default:
					break;

				}
			}

		} catch (Exception e) { // NOPMD
		// Ignore exception
		}

	}





	/**
	 * Last action of the robot.
	 */

	public String manageException(String action, Exception exception) throws Exception {
	// We get the message of the exception
		String errorMessage = ExceptionUtils.getRootCause(exception).getMessage();
		// We send a screenshot to the log so the user can see the screen in the moment
		// of the error
		// This is a very useful thing to do
		server.sendScreen("Screenshot at the moment of the error");
		server.setCurrentItemResultToWarn(exception.getCause().getLocalizedMessage());
		// If we have a FatalException we should abort the execution.
		if (ExceptionUtils.indexOfThrowable(exception, JidokaItemException.class) >= 0) {

			server.error(StringUtils.isBlank(errorMessage) ? "Item error" : errorMessage);
			return IRobot.super.manageException(action, exception);
		}
		else if(ExceptionUtils.indexOfThrowable(exception, JidokaQueueException.class) >= 0){
			server.error(StringUtils.isBlank(errorMessage) ? "Queue error" : errorMessage);
			return IRobot.super.manageException(action, exception);

		}
		server.warn("Unknown exception!");

		// If we have any other exception we must abort the execution, we don't know
		// what has happened

		return IRobot.super.manageException(action, exception);
	}
	public void end()  {
		server.info("End of Process");
	}
	public void customerEnd()  {
		server.info("End of Customer Web");
	}
	public void googleEnd()  {
		server.info("End of Google Doc Web");
	}
	
}
