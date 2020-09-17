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

	private Excel_Input currentItem;

	private static final int FIRST_ROW = 0;

	private IQueueItem currentItemQueue;
	private int currentItemIndex;

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
	public boolean start() throws Exception {
		qmanager = server.getQueueManager();
		queueCommons = new QueueCommons();
		excelDSRow = new ExcelDSRow();
		queueCommons.init(qmanager);
		dataProvider = IJidokaDataProvider.getInstance(this, IJidokaDataProvider.Provider.EXCEL);
		server.setNumberOfItems(1);
		excel = IExcel.getExcelInstance(this);
		server.setNumberOfItems(1);
		server = (IJidokaServer< ? >) JidokaFactory.getServer();

		client = IClient.getInstance(this);

		browser = IWebBrowserSupport.getInstance(this, client);
		return IRobot.super.startUp();
	}



	/**
	 * Open Web Browser
	 * @throws Exception
	 */
	public void openBrowser() throws Exception  {

		browserType = server.getParameters().get("Browser");
		
		// Select browser type
		if (StringUtils.isBlank(browserType)) {
			server.info("Browser parameter not present. Using the default browser CHROME");
			browser.setBrowserType(EBrowsers.CHROME);
			browserType = EBrowsers.CHROME.name();
		} else {
			EBrowsers selectedBrowser = EBrowsers.valueOf(browserType);
			browserType = selectedBrowser.name();
			browser.setBrowserType(selectedBrowser);
			server.info("Browser selected: " + selectedBrowser.name());
		}
		
		// Set timeout to 60 seconds
		browser.setTimeoutSeconds(60);

		// Init the browser module
		browser.initBrowser();

		//This command is uses to make visible in the desktop the page (IExplore issue)
		if (EBrowsers.INTERNET_EXPLORER.name().equals(browserType)) {
			client.clickOnCenter();
			client.pause(3000);
		}

	}

	/**
	 * Navigate to Web Page
	 * 
	 * @throws Exception
	 */
	public void navigateToWeb() throws Exception  {
		
		server.setCurrentItem(1, HOME_URL);
		
		// Navegate to HOME_URL address
		browser.navigate(HOME_URL);

		// we save the screenshot, it can be viewed in robot execution trace page on the console
		server.sendScreen("Screen after load page: " + HOME_URL);
		
		server.setCurrentItemResultToOK("Success");
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
		Path inputFile = Paths.get(server.getCurrentDir(), inputFilePath);


		dataProvider = IJidokaDataProvider.getInstance(this, IJidokaDataProvider.Provider.EXCEL);
		dataProvider.init(String.valueOf(inputFile), null, FIRST_ROW, new Excel_Input_RowMapper());

		try {

			// Get the next row, each row is a item
			while (dataProvider.nextRow()) {

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
		server.info("End process");
	}
	
}
