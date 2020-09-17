package com.novayre.jidoka.robot.test;

import com.novayre.jidoka.client.api.queue.IQueueManager;
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

	private IJidokaExcelDataProvider<ExcelDSRow> dataProvider;

	private IExcel excel;

	private IQueueManager qmanager;

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
	public boolean start() throws Exception {
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
		return IRobot.super.startUp();
	}


/**
	 * Navigate to Web Page
	 * 
	 * @throws Exception
	 */
	public void navigateToCustomerWeb() throws Exception  {
		//Set File Name for Customer Xpath
		String cXpathFileName = server.getEnvironmentVariables().get("customerXpathFileName").toString();
		webApplication.PerformOperation(cXpathFileName);

	}
	public void navigateToGoogleWeb() throws Exception  {
		//Set File Name for Google Xpath
		String gXpathFileName = server.getEnvironmentVariables().get("GoogleXpathFileName").toString();
		webApplication.PerformOperation(gXpathFileName);

	}
	public void customerRetry() throws Exception{
		returnType = webApplication.RetryRequired();
	if (returnType.contains("Yes")){
		navigateToCustomerWeb();
	}
	else if(returnType.contains("No")){
		customerEnd();
	}

	}
	public void googleRetry() throws Exception{
		returnType = webApplication.RetryRequired();
		if (returnType.contains("Yes")){
			navigateToCustomerWeb();
		}
		else{
			googleEnd();
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
