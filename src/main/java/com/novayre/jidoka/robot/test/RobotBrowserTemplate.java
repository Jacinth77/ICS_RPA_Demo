package com.novayre.jidoka.robot.test;

import com.novayre.jidoka.client.api.exceptions.JidokaQueueException;
import com.novayre.jidoka.client.api.queue.*;
import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IJidokaDataProvider;
import com.novayre.jidoka.data.provider.api.IJidokaExcelDataProvider;
import org.apache.commons.lang.StringUtils;

import com.novayre.jidoka.browser.api.EBrowsers;
import com.novayre.jidoka.browser.api.IWebBrowserSupport;
import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.multios.IClient;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

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
	private Excel_Input currentItem;

	private static final int FIRST_ROW = 0;

	private IQueueItem currentItemQueue;
	private int currentItemIndex;

	private ICS_WebApplication webApplication;

	private String returnType ="No";

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

		return IRobot.super.startUp();

	}
	
	/**
	 * Action "start".
	 */
	public void start() throws Exception {
		server = (IJidokaServer< ? >) JidokaFactory.getServer();

		client = IClient.getInstance(this);

		browser = IWebBrowserSupport.getInstance(this, client);
		qmanager = server.getQueueManager();
		queueCommons = new QueueCommons();
		webApplication =new ICS_WebApplication();
		excelDSRow = new ExcelDSRow();
		queueCommons.init(qmanager);
		dataProvider = IJidokaDataProvider.getInstance(this, IJidokaDataProvider.Provider.EXCEL);
		server.setNumberOfItems(1);
		excel = IExcel.getExcelInstance(this);
		server.setNumberOfItems(1);
		server = (IJidokaServer< ? >) JidokaFactory.getServer();

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
		if (returnType.contains("maxCountReached")){
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
		server.info("InputID "+currentItem.getInput_ID());
		webApplication.PerformOperation(cXpathFileName,currentItem.getInput_ID());

	}
	public void navigateToGoogleWeb() throws Exception  {
		//Set File Name for Google Xpath
		String gXpathFileName = server.getEnvironmentVariables().get("GoogleXpathFileName").toString();
		server.info("InputID "+currentItem.getInput_ID());
		webApplication.PerformOperation(gXpathFileName,currentItem.getInput_ID());

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

		if (StringUtils.isNotBlank(qmanager.preselectedQueue())) {

			selectedQueueID = qmanager.preselectedQueue();
			server.info("Selected queue ID: " + selectedQueueID);
			currentQueue = queueCommons.getQueueFromId(selectedQueueID);
		} else {

			String inputFilePath=server.getEnvironmentVariables().get("InputFilePath").toString();
			selectedQueueID = queueCommons.createQueue(inputFilePath);
			server.info("Queue ID: " + selectedQueueID);
			addItemsToQueue();


			currentQueue = queueCommons.getQueueFromId(selectedQueueID);

		}
	}

	private void addItemsToQueue() throws Exception {


		String inputFilePath=server.getEnvironmentVariables().get("InputFilePath").toString();
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
			}

		} catch (Exception e) {
			throw new JidokaQueueException(e);
		} finally {

			try {
				// Close the excel file
				dataProvider.close();
			} catch (IOException e) {
				throw new JidokaQueueException(e);
			}
		}
	}

	public String hasMoreItems() throws Exception {

		// retrieve the next item in the queue
		//QueueCommons queueCommons = new QueueCommons();
		currentItemQueue = queueCommons.getNextItem(currentQueue);

		if (currentItemQueue != null) {



			currentItem= new Excel_Input();

			currentItem.setInput_ID(currentItemQueue.functionalData().get("Input_ID"));
			currentItem.setInput_ID(currentItemQueue.functionalData().get("Status"));

			server.info(currentItem.getInput_ID());

			// set the stats for the current item

			server.setCurrentItem(currentItemIndex++, currentItemQueue.key());

			return "Yes";
		}

		return "No";
	}

	public void releaseitems() throws IOException, JidokaQueueException {

		Map<String, String> funcData = currentItemQueue.functionalData();
		funcData.put(Excel_Input_RowMapper.Status,"Success");

		ReleaseItemWithOptionalParameters rip = new ReleaseItemWithOptionalParameters();
		rip.functionalData(funcData);
		qmanager.releaseItem(rip);

	}

	public void write_to_Excel(){

	}
	public void Move_File(){

	}
	public void closeQueue(){

	}
	public void getdata(){

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
