/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package createusers;
import java.awt.EventQueue;
import javax.swing.JFrame;
import javax.swing.JButton;
import javax.swing.JTextPane;


import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.AbstractAction;
import javax.swing.AbstractButton;

import java.awt.event.ActionEvent;
import javax.swing.Action;
import java.awt.event.ActionListener;
import javax.swing.SwingConstants;

import org.apache.commons.lang3.RandomStringUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import javax.swing.JTextArea;
import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateUsers {
    private static String driverpath = "G://Work//Netbean project//CreateUsers//drivers//chromedriver.exe";
    private JFrame frame;
    private static JTextField txtUrl;
    private final Action action = new SwingAction();
    private static JButton btnStop;
    private static AbstractButton btnNewButton;
    
    private static JButton btnLoadData;    
    MyClass myClass = new MyClass();
    ReadXLSXClass readXLSXClass;
    private static String filepath = "test.xlsx";
    public static String url;

    public static int user = 0;
    private JLabel lblNewLabel;
    private JLabel lblUsername;
    private static JTextArea textPane;
    private static JTextPane textPane_1;
    private static JTextPane textPane_2;
    private static JTextPane textPane_3;
    public static WebDriver driver;
    static WebDriverWait wait;
    static boolean exit = false;
    static StringBuilder nr = new StringBuilder();
    static StringBuilder usr = new StringBuilder();
    static StringBuilder pwd = new StringBuilder();
    static StringBuilder status = new StringBuilder();
    static String[][] info;
    
    /*static Logger logger = LoggerFactory.getLogger(UserGUI.class);*/
    
    public static String randomizer(){

        return RandomStringUtils.randomAlphabetic(5);
    }

    public static String randomizerInt(){
        return RandomStringUtils.randomNumeric(7);
    }
    
    public static class MyClass extends Thread {   
    	public void run() {
            while(!exit){
                user++;	
//		logger.debug("Set the system property for chrome driver");
                //System.setProperty("webdriver.chrome.driver", "drivers//chromedriver.exe");

                System.setProperty("webdriver.chrome.driver", driverpath);

//		logger.debug("Instantiate ChromeDriver");
                driver = new ChromeDriver();

//			    logger.debug("Maximize the window");
                driver.manage().window().maximize();

//			    logger.debug("Open the web page");
                url = txtUrl.getText();
                try{
                      driver.get(url);
                }catch(WebDriverException e){
                    JOptionPane.showMessageDialog(null, "Cannot navigate to the provided URL\nURL must have the following structure : https://www.google.ro");
                    driver.quit();
                    btnNewButton.setEnabled(true);
                    btnLoadData.setEnabled(true);
                    btnStop.setEnabled(false);
                    btnStop.setText("Stop");
                }

                wait = new WebDriverWait(driver, 10);

                WebElement firstName = driver.findElement(By.id("firstName"));
//			    logger.debug("id is :" + firstName.toString());
                //firstName.click();
                firstName.sendKeys(randomizer());

                WebElement middleName = driver.findElement(By.id("lastName"));
               /* new Actions(driver).moveToElement(middleName).click().perform();*/
                middleName.sendKeys(randomizer());
                //middleName.sendKeys(Keys.TAB);

                WebElement logonId = driver.findElement(By.xpath("//div[@id=\"login-email\"]//input[@id=\"logonId\"]"));
                //WebElement logonId = driver.findElement(By.name("logonId"));
                //new Actions(driver).moveToElement(logonId).perform();
                //logonId.click();
                String logonIdString = randomizer() + "@gmail.com";
                logonId.sendKeys(logonIdString);
                //logonId.sendKeys(Keys.TAB);
                try {
                    Thread.sleep(1000);  
                } catch (InterruptedException ex) {
                    Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
                }


                WebElement confirmLogonId = driver.findElement(By.id("confirmLogonId"));
                confirmLogonId.sendKeys(logonIdString);
                confirmLogonId.sendKeys(Keys.TAB);

                WebElement logonPassword = driver.findElement(By.id("logonPassword"));
                String loginPassword = randomizer()+"Rd1m";
                logonPassword.sendKeys(loginPassword);
                logonPassword.sendKeys(Keys.TAB);
                WebElement logonPasswordVerify = driver.findElement(By.id("logonPasswordVerify"));
                logonPasswordVerify.sendKeys(loginPassword);
                logonPasswordVerify.sendKeys(Keys.TAB);

                WebElement WC_PersonalInfoExtension_birth_date = driver.findElement(By.id("WC_PersonalInfoExtension_birth_date"));
                Select day = new Select(WC_PersonalInfoExtension_birth_date);
                day.selectByValue("1");

                WebElement WC_PersonalInfoExtension_birth_month = driver.findElement(By.id("WC_PersonalInfoExtension_birth_month"));
                Select month = new Select(WC_PersonalInfoExtension_birth_month);
                month.selectByValue("1");

                WebElement WC_PersonalInfoExtension_birth_year = driver.findElement(By.id("WC_PersonalInfoExtension_birth_year"));
                Select year = new Select(WC_PersonalInfoExtension_birth_year);
                year.selectByValue("1991");

                WebElement male = driver.findElement(By.id("male"));
                male.click();

                WebElement country = driver.findElement(By.id("country"));
                Select countrySelect = new Select(country);
                countrySelect.selectByValue("AU");

                WebElement address = driver.findElement(By.id("address"));
                address.sendKeys("8 Central Ave, Eveleigh NSW 2015");
                address.sendKeys(Keys.SPACE);
                try {
                    Thread.sleep(5000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                address.sendKeys(Keys.ARROW_DOWN);
                address.sendKeys(Keys.ENTER);

                WebElement phoneType1 = driver.findElement(By.id("phoneType1"));
                Select phoneType = new Select(phoneType1);
                phoneType.selectByValue("CEL");

                WebElement phone1 = driver.findElement(By.id("phone1"));
                phone1.sendKeys("614" + randomizerInt());
                //phone1.click();
                //logonId.sendKeys(Keys.TAB);
                try {
                    Thread.sleep(1000);
                } catch (InterruptedException ex) {
                    Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
                }
                driver.findElement(By.xpath("//input[@type='radio' and @id='myerone-member-no']")).click();

             //   uncomment this to click on create button
                driver.findElement(By.id("create-account")).click();

                //Check the result if the account create.
                WebElement result = driver.findElement(By.id("ErrorMessageText"));
                System.out.println(result.getText());

                String status_txt = "FAILED";
                if(result.getText().contains("Successfully"))
                    status_txt = "DONE";
                if(result.getText().contains("already exists"))
                    status_txt = "EXIST";
                    //driver.quit();
                nr.append(String.valueOf(user)+"\n");
                textPane.setText(nr.toString());


                usr.append(logonIdString+"\n");
                textPane_1.setText(usr.toString());

                pwd.append(loginPassword+"\n");
                textPane_2.setText(pwd.toString());

                exit=true;                
                status.append(status_txt);
                status.append("\n");
                textPane_3.setText(status.toString());
                

                btnNewButton.setEnabled(true);
                btnLoadData.setEnabled(true);
                btnStop.setEnabled(false);
                btnStop.setText("Stop");
						
            }

    	}
    }

	/**
	 * Launch the application.
	 */
    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {                    
                    CreateUsers window = new CreateUsers();
                    window.frame.setVisible(true);                    
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * Create the application.
     */
    public CreateUsers() {
            initialize();
    }
    
    /**
     * Read Excel Info
     */
    public static class ReadXLSXClass extends Thread {   
        private String filepath;
        ReadXLSXClass(String path) {
            this.filepath = path;
        } 
        
    	public void run() {            
            System.setProperty("webdriver.chrome.driver", driverpath);
            
            InputStream ExcelFileToRead = null;
            try {
                ExcelFileToRead = new FileInputStream(filepath);
            } catch (FileNotFoundException ex) {
                Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
                System.out.println("Excel File does not exist!");
                return;
            }


            /**Read Excel**/
            XSSFWorkbook  wb = null;
            try {
                wb = new XSSFWorkbook(ExcelFileToRead);
            } catch (IOException ex) {
                //Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
                System.out.println("Excel File does not exist!");
            }                       

            XSSFWorkbook test = new XSSFWorkbook();       

            XSSFSheet sheet = wb.getSheetAt(0);            

            XSSFRow row; 
            XSSFCell cell;

            Iterator rows = sheet.rowIterator();
            info = new String[sheet.getLastRowNum() + 1][12];

            int _crow = -1, _ccell = 0;

            try {
                while (rows.hasNext())
                {
                        _crow ++;
                        row=(XSSFRow) rows.next();
                        Iterator cells = row.cellIterator();
                        _ccell = -1;
                        while (cells.hasNext())
                        {
                            _ccell ++;
                            cell=(XSSFCell) cells.next();
                            
                            if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING)
                            {
                                info[_crow][_ccell] = cell.getStringCellValue();                                    
                            }
                            else if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
                            {
                                Double ddf = new Double(cell.getNumericCellValue());
                                
                                info[_crow][_ccell] = String.format("%d", ddf.intValue());
                            }
                        }                  
                }
            } catch(Exception fe) {
                fe.printStackTrace();
                System.out.println("Error excel type");
                return;
            }

            // check excel type
            try {
                String[] header = {"First Name","Last Name","Login Email Address","Confirm Login Email Address","Choose Password","Confirm Password","DOB","Gender","Country","Address","Contact Number","Myer Card"};
                for(int i=0; i<sheet.getLastRowNum(); i++) {
                    if(!info[0][i].equals(header[i]))
                    {
                        System.out.println("Error excel type");
                        return;
                    }
                }
            } catch(Exception ee) {
                System.out.println("Error excel type");
                return;
            }

            for(int i=1; i<info.length;i++) {                
                createUser(info[i]);
            }

            //upload create
            try {
                test.close();
                wb.close();
                ExcelFileToRead.close();
            } catch (IOException ex) {
                Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
            }
                
            exit = true;
           
            
    	}
    }
    
    private static void createUser(String[] info) {
          
        user++;	
//	logger.debug("Instantiate ChromeDriver");
        driver = new ChromeDriver();

//	logger.debug("Maximize the window");
        driver.manage().window().maximize();

//			    logger.debug("Open the web page");
        url = txtUrl.getText();
        try{
            driver.get(url);
        }catch(WebDriverException e){
            JOptionPane.showMessageDialog(null, "Cannot navigate to the provided URL\nURL must have the following structure : https://www.google.ro");
            driver.quit();
            btnNewButton.setEnabled(true);
            btnLoadData.setEnabled(true);
            btnStop.setEnabled(false);
            btnStop.setText("Stop");
        }

        wait = new WebDriverWait(driver, 10);

        WebElement firstName = driver.findElement(By.id("firstName"));
        firstName.sendKeys(info[0]);   

        WebElement middleName = driver.findElement(By.id("lastName"));
        middleName.sendKeys(info[1]);

        WebElement logonId = driver.findElement(By.xpath("//div[@id=\"login-email\"]//input[@id=\"logonId\"]"));                        
        logonId.sendKeys(info[2]);

        try {
            Thread.sleep(1000);  
        } catch (InterruptedException ex) {
            Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
        }


        WebElement confirmLogonId = driver.findElement(By.id("confirmLogonId"));
        confirmLogonId.sendKeys(info[3]);
        confirmLogonId.sendKeys(Keys.TAB);

        WebElement logonPassword = driver.findElement(By.id("logonPassword"));
        String loginPassword = randomizer()+"Rd1m";
        logonPassword.sendKeys(info[4]);
        logonPassword.sendKeys(Keys.TAB);


        WebElement logonPasswordVerify = driver.findElement(By.id("logonPasswordVerify"));
        logonPasswordVerify.sendKeys(info[5]);
        logonPasswordVerify.sendKeys(Keys.TAB);

        String year1;
        String month1;
        String day1;
        try {
            String[] dd = info[6].split("/");
            year1 = dd[2];
            month1 = dd[1];
            day1 = dd[0];
        } catch(Exception e){
            year1 = "1900";
            month1 = "1";
            day1 = "1";
        }

        WebElement WC_PersonalInfoExtension_birth_date = driver.findElement(By.id("WC_PersonalInfoExtension_birth_date"));
        Select day = new Select(WC_PersonalInfoExtension_birth_date);
        day.selectByValue(day1);

        WebElement WC_PersonalInfoExtension_birth_month = driver.findElement(By.id("WC_PersonalInfoExtension_birth_month"));
        Select month = new Select(WC_PersonalInfoExtension_birth_month);
        month.selectByValue(month1);

        WebElement WC_PersonalInfoExtension_birth_year = driver.findElement(By.id("WC_PersonalInfoExtension_birth_year"));
        Select year = new Select(WC_PersonalInfoExtension_birth_year);
        year.selectByValue(year1);

        if(info[7].equals("M")) {
            WebElement male = driver.findElement(By.id("male"));
            male.click();
        } else {
            WebElement female = driver.findElement(By.id("female"));
            female.click();
        }

        WebElement country = driver.findElement(By.id("country"));
        Select countrySelect = new Select(country);
        countrySelect.selectByValue(info[8]);

        WebElement address = driver.findElement(By.id("address"));
        address.sendKeys(info[9]);
        address.sendKeys(Keys.SPACE);

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        address.sendKeys(Keys.ARROW_DOWN);
        address.sendKeys(Keys.ENTER);

        WebElement phoneType1 = driver.findElement(By.id("phoneType1"));
        Select phoneType = new Select(phoneType1);
        phoneType.selectByValue("CEL");

        WebElement phone1 = driver.findElement(By.id("phone1"));
        phone1.sendKeys(info[10]);

        try {
            Thread.sleep(1000);
        } catch (InterruptedException ex) {
            Logger.getLogger(CreateUsers.class.getName()).log(Level.SEVERE, null, ex);
        }
        driver.findElement(By.xpath("//input[@type='radio' and @id='myerone-member-no']")).click();

     //   uncomment this to click on create button
        driver.findElement(By.id("create-account")).click();

        //Check the result if the account create.
        WebElement result = driver.findElement(By.id("ErrorMessageText"));
        System.out.println(result.getText());

        String status_txt = "FAILED";
        if(result.getText().contains("Successfully"))
            status_txt = "DONE";
        if(result.getText().contains("already exists"))
            status_txt = "EXIST";
            //driver.quit();
        nr.append(String.valueOf(user)+"\n");
        textPane.setText(nr.toString());


        usr.append(info[2]+"\n");
        textPane_1.setText(usr.toString());

        pwd.append(info[4]+"\n");
        textPane_2.setText(pwd.toString());

        exit=true;
        status.append(status_txt);
        status.append("\n");
        textPane_3.setText(status.toString());

        btnNewButton.setEnabled(true);
        btnLoadData.setEnabled(true);
        btnStop.setEnabled(false);
        btnStop.setText("Stop");

        driver.close();
        
    }
    
    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 544, 334);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(null);

        btnNewButton = new JButton("Start");
        btnNewButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent arg0) {
                myClass.start();
                btnNewButton.setEnabled(false);
                btnLoadData.setEnabled(false);
                btnStop.setEnabled(true);

            }
        });
        btnNewButton.setAction(action);
        btnNewButton.setBounds(10, 11, 120, 44);
        frame.getContentPane().add(btnNewButton);

        btnStop = new JButton("Stop");
        btnStop.setEnabled(false);
        
        btnStop.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {                
                exit = true;                 
                btnNewButton.setEnabled(false);
                btnLoadData.setEnabled(false);
                btnStop.setEnabled(false);
                
                btnStop.setText("Stopping...");
                JOptionPane.showMessageDialog(null, "Please wait till execution of thread stops");
            }
        });
        btnStop.setBounds(140, 11, 120, 44);
        frame.getContentPane().add(btnStop);

        btnLoadData = new JButton("Load Data");
        btnLoadData.setEnabled(true);
        btnLoadData.setBounds(270, 11, 120, 44);
        frame.getContentPane().add(btnLoadData);

        btnLoadData.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                readXLSXClass = new ReadXLSXClass(filepath);
                exit = true; 
                
                btnLoadData.setEnabled(false);
                btnNewButton.setEnabled(false);
                btnStop.setEnabled(false);
                btnStop.setText("Loading...");
                btnNewButton.setEnabled(false);                
                readXLSXClass.start();
            }
        });

        JButton btnLoadProxy = new JButton("Load Proxy");
        btnLoadProxy.setEnabled(false);
        btnLoadProxy.setBounds(400, 11, 120, 44);
        frame.getContentPane().add(btnLoadProxy);

        textPane = new JTextArea();
        textPane.setEditable(false);
        textPane.setBounds(10, 91, 34, 140);
        frame.getContentPane().add(textPane);


        textPane_1 = new JTextPane();
        textPane_1.setEditable(false);
        textPane_1.setBounds(54, 91, 185, 140);
        frame.getContentPane().add(textPane_1);

        textPane_2 = new JTextPane();
        textPane_2.setEditable(false);
        textPane_2.setBounds(249, 91, 189, 140);
        frame.getContentPane().add(textPane_2);

        textPane_3 = new JTextPane();
        textPane_3.setEditable(false);
        textPane_3.setForeground(Color.GREEN);
        textPane_3.setBounds(448, 91, 72, 140);
        frame.getContentPane().add(textPane_3);

        lblNewLabel = new JLabel("Nr");
        lblNewLabel.setBounds(20, 66, 68, 14);
        frame.getContentPane().add(lblNewLabel);


        lblUsername = new JLabel("Email");
        lblUsername.setBounds(113, 66, 68, 14);
        frame.getContentPane().add(lblUsername);

        JLabel lblPassword = new JLabel("Password");
        lblPassword.setBounds(306, 66, 68, 14);
        frame.getContentPane().add(lblPassword);

        JLabel lblProxy = new JLabel("Status");
        lblProxy.setBounds(460, 66, 46, 14);
        frame.getContentPane().add(lblProxy);

        txtUrl = new JTextField();
        txtUrl.setHorizontalAlignment(SwingConstants.LEFT);
        txtUrl.setText("https://www.myer.com.au/webapp/wcs/stores/servlet/UserRegistrationForm?new=Y&catalogId=10051&myAcctMain=1&langId=-1&storeId=10251");
        url = txtUrl.getText();

        txtUrl.setBounds(10, 258, 510, 26);
        frame.getContentPane().add(txtUrl);
        txtUrl.setColumns(10);

        JLabel lblUrl = new JLabel("URL");
        lblUrl.setBounds(245, 242, 68, 14);
        frame.getContentPane().add(lblUrl);
    }
    private class SwingAction extends AbstractAction {
        public SwingAction() {
                putValue(NAME, "Start");
                putValue(SHORT_DESCRIPTION, "Some short description");
        }
        public void actionPerformed(ActionEvent e) {
        }
    }
}
