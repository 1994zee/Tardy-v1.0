using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tardy.model;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;

namespace Tardy
{
    class Program
    {
        public static void ErrorLogging1(Exception ex)
        {
            try
            {
                string strPath = "Log.txt"; //name of the text file to be created
                if (!File.Exists(strPath))//checking whether a file with above name already exists or not
                {
                    File.Create(strPath).Dispose();//creates a new file if doesn't exists earlier.
                }
                using (StreamWriter sw = File.AppendText(strPath))//opening stream to write into the file
                {
                    //below code writes Exception details in the text file created above
                    sw.WriteLine("=============Error Logging ===========");
                    sw.WriteLine("===========Start============= " + DateTime.Now);
                    sw.WriteLine("Error Message: " + ex.Message);
                    sw.WriteLine("Stack Trace: " + ex.StackTrace);
                    sw.WriteLine("===========End============= " + DateTime.Now);
                }

                //below line of code sends the above created text file as an attachment to the desired email
                // First parameter: from email address
                //second parameter: to email address
                //third parameter: Email subject
                // Fourth parameter: email body
                //fifth parameter: local address of the above created text file
                Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "ba@lightsourcehr.com", "Error log", ex.Message, strPath, "data.txt");
            }
            catch(Exception exi)
            {
                Console.WriteLine(exi);
            }
        }
        static void Main(string[] args)
        {
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            TimeSpan waitTime = new TimeSpan(0, 1, 0);// varaiable for time wait of one minute
            List<Record> withEmployeeID = new List<Record>();
            List<Record> withoutEmployeeID = new List<Record>();
            List<Record> records = new List<Record>();
            List<Record_2> records_2 = new List<Record_2>();
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
           Wait:
            Console.ForegroundColor = ConsoleColor.Magenta; //setting console text color
            Console.WriteLine("System on wait");
            Console.ForegroundColor = ConsoleColor.White; //setting console text color
            while (true) //infinited loop
            {
                //getting current time and data
                String wait = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                if (wait == "10")
                {
                    break; //ends the wait condition
                }
                System.Threading.Thread.Sleep(waitTime); //puts system to sleep for one minute
            }
            IWebDriver gc = new ChromeDriver(options);
            //Client space login process
            try
            {
                //redirects browser to clientspace url
                // Fills up the user id field
                //fills up the password field.
                //press enter to proceed for login authentication
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
                gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
                gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
                gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Client space login failed");
                Console.ForegroundColor = ConsoleColor.White;
            }
            // End

            //opening report WCLNF1 Ancillary Risk Fees (LightBot Admins Only)
            try
            {
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\WCLNF1+Ancillary+Risk+Fees");
                //System.Threading.Thread.Sleep(new TimeSpan(0, 0, 30));
                //gc.FindElement(By.Id("ndbfc0")).Click();
                //gc.FindElements(By.TagName("option"))[3].Click();
                //gc.FindElement(By.Id("updateBtnP")).Click();
                System.Threading.Thread.Sleep(new TimeSpan(0, 1, 0));
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Loading report Data failed");
                Console.ForegroundColor = ConsoleColor.White;
            }
            // End

            //extracting data from the report
            // code for extracting data from the report and loading it into application memory
            try
            {
                //data is loaded on a table strucutre in the html of the source webpage
                // table contains two alernating type of rows
                // one type named as "alternatingItems" in the html
                //second type named as "ReportItem" in the html
                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                //below is a for loop which finds all the rows named as "alternating item"
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    //find all the colums of the row
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    //load them in the records class object by calling an overloaded constructor
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text
                        , cols.ElementAt(8).Text, cols.ElementAt(9).Text, cols.ElementAt(10).Text, cols.ElementAt(11).Text);
                    //adds the above created object to the list
                    records.Add(r);

                }
                //getting records from normal rows

                //getting reports from alteranating rows
                //below is a for loop which finds all the rows named as "Report Item"
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    //find all the colums of the row
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    //load them in the records class object by calling an overloaded constructor
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text
                        , cols.ElementAt(8).Text, cols.ElementAt(9).Text, cols.ElementAt(10).Text, cols.ElementAt(11).Text);
                    //adds the above created object to the list
                    records.Add(r);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Data Extraction Completed !");
                Console.ForegroundColor = ConsoleColor.White;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Data scrapping failed !");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //..extraction of data ends here

            //extracting data from second report
            try
            {
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\WCLNF2+Client+Fees");
                System.Threading.Thread.Sleep(new TimeSpan(0, 0, 30));
                //gc.FindElement(By.Id("ndbfc0")).Click();
                //gc.FindElements(By.TagName("option"))[3].Click();
                //gc.FindElement(By.Id("updateBtnP")).Click();
            //    System.Threading.Thread.Sleep(new TimeSpan(0, 1, 0));
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Loading report Data failed");
                Console.ForegroundColor = ConsoleColor.White;
            }
            // End

            //extracting data from the report
            // code for extracting data from the report and loading it into application memory
            try
            {
                //data is loaded on a table strucutre in the html of the source webpage
                // table contains two alernating type of rows
                // one type named as "alternatingItems" in the html
                //second type named as "ReportItem" in the html
                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                //below is a for loop which finds all the rows named as "alternating item"
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    //find all the colums of the row
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    //load them in the records class object by calling an overloaded constructor
                    Record_2 r = new Record_2(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text
                        , cols.ElementAt(8).Text);
                    //adds the above created object to the list
                    records_2.Add(r);

                }
                //getting records from normal rows

                //getting reports from alteranating rows
                //below is a for loop which finds all the rows named as "Report Item"
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    //find all the colums of the row
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    //load them in the records class object by calling an overloaded constructor
                    Record_2 r = new Record_2(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text
                        , cols.ElementAt(8).Text);
                    //adds the above created object to the list
                    records_2.Add(r);
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Data Extraction Completed !");
                Console.ForegroundColor = ConsoleColor.White;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Data scrapping failed !");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //...................................................









            List<Record> remove = new List<Record>();
            //PRINTING EXTRACTED DATA FOR TESTING PURPOSE
            if(records.Count>0)
            {
                foreach(Record r in records)
                {
                    if(r.days_since_reported<2)
                    {
                        remove.Add(r);
                    }
                    Console.WriteLine(r.cs_claim_number+"-"+r.days_since_reported);
                }
                foreach(Record r in remove)
                {
                    records.Remove(r);
                }
                Console.WriteLine("------");
                foreach (Record r in records)
                {   
                    Console.WriteLine(r.cs_claim_number);
                }
            }
            else
            {
                Console.WriteLine("No data extracted");
            }
            remove.Clear();

            //filtering data
            //over employee id column
            foreach(Record r in records)
            {
                if(r.employee_id!="")
                {
                    withEmployeeID.Add(r);
                }
                else
                {
                    withoutEmployeeID.Add(r);
                }
            }
            records.Clear();
            //...
            List<Record> importable = new List<Record>();
            importable.AddRange(withEmployeeID);
            List<Record> cases = new List<Record>();
            foreach(Record r in withoutEmployeeID)
            {
                if(r.injured_employee!="" && r.location!="")
                {
                    importable.Add(r);
                }
                else
                {
                    cases.Add(r);
                }
            }
            withEmployeeID.Clear();
            withoutEmployeeID.Clear();
            //creating cases step
            if (cases.Count < 0)
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Creating cases for " + cases.Count + " number of records.");
                Console.ForegroundColor = ConsoleColor.Blue;
                int i = 1;
                foreach (Record record in cases)
                {
                    Console.WriteLine("\nCreating case # " + i);
                    i++;
                    try
                    {
                        gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/peo/client");
                        gc.FindElement(By.Id("dropdownMenu1")).Click();
                        gc.FindElement(By.Id("All")).Click();
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ClientNumber")).SendKeys(record.client_id.ToString()); // put the client number here.
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.ClassName("formSearchBtn")).Click();
                        System.Threading.Thread.Sleep(1000);
                        gc.FindElements(By.ClassName("cs-underline"))[1].Click();
                        gc.FindElement(By.XPath("//*[@id='customHeaderHtml']/div[2]/div[6]/div/div[1]/table/tbody/tr/td[1]/span")).Click();
                        gc.FindElement(By.XPath("//*[@id='lstDataform_btnAdd']")).Click();
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys("R");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys(Keys.Tab);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys("M");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Enter);
                        gc.FindElement(By.XPath("//*[@id='CallerName']")).SendKeys("Lightbot");
                        gc.FindElement(By.XPath("//*[@id='EmailAddress']")).SendKeys("lightbot@lightsourcehr.com");
                        DateTime dateTime = DateTime.Now;
                        var date = DateTime.Now.ToString("MM/dd/yyyy");
                        gc.FindElement(By.XPath("//*[@id='DueDate']")).SendKeys(date);

                        System.Threading.Thread.Sleep(10000);
                        if (record.carrier_claim_number == "")
                        {
                            gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys("#"+record.cs_claim_number.ToString());
                        }
                        else
                        {
                            gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys("#"+record.carrier_claim_number.ToString());
                        }
                        
                        gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Result : Case creation Sucessfull.");

                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Result : Case creation failed..!");
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }
            }


            int bill_rate = 0;
            
            Console.Clear();
            List<Record> final = new List<Record>();
            foreach(Record r in importable)
            {
                
                foreach(Record_2 i in records_2.Where(s=>s.client_id==r.client_id))
                {
                    if(i.surcharge_type== "Late Claims Reporting Fee Addtl Days")
                    {
                        i.bill_rates = i.bill_rates.Replace(".00", "").Replace(",", "");
                        bill_rate += (Convert.ToInt32(i.bill_rates) * Convert.ToInt32(r.days_since_reported));
                        Console.ForegroundColor = ConsoleColor.Blue;
                    }
                    else
                    {
                        i.bill_rates = i.bill_rates.Replace(".00", "").Replace(",", "");
                        bill_rate += Convert.ToInt32(i.bill_rates);
                        Console.ForegroundColor = ConsoleColor.Cyan;
                    }
                    Console.WriteLine("--"+i.bill_rates);
                }
                r.bill_rate = bill_rate.ToString();
                bill_rate = 0;
                final.Add(r);
                Console.WriteLine(r.days_since_reported+"---"+r.bill_rate+"___"+r.doi);
            }
            importable.Clear();
            string delimiter = "\t";
            if (final.Count > 0)
            {
                try
                {
                    using (var writer = new System.IO.StreamWriter("data.txt"))
                    {
                        foreach (Record i in final)
                        {
                            
                            writer.WriteLine(i.bill_date + delimiter +i.bill_event_code + delimiter +i.bill_rate + delimiter +i.bill_unit + delimiter +i.client_id
                                + delimiter +i.employee_id + delimiter +i.location + delimiter+ delimiter +i.comment+i.doi);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Import file creation failed.");
                }
            }

            

            //case Creation Ends here

            //Importing data into prism
            // imports the valid data from the list with valid location into a text file
            //which is than imported onto prism for the completion of process
            if (final.Count > 0) //checking whether valid records were found or not
            {
                

                //logging into prism
                top:
                //redirecting brower to prismhr web page
                gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                System.Threading.Thread.Sleep(1000);
                //logging in prism
                try
                {
                    //The browser will be on login page
                    //enters the user name into the user name field
                    gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                    //enters the password into the password field
                    gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
                    //press sign in button
                    gc.FindElement(By.XPath("//*[@id='button8v1']")).Click();
                    //puts system to wait for one second
                    System.Threading.Thread.Sleep(new TimeSpan(0,0,20));
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Prism login failed!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Exception e = new Exception("Prism Login failed !");
                //    ErrorLogging(ex);
                  //  ErrorLogging(e);
                  //  mail.Add("Prism Login failed !");
                }
                //after login
                try
                {
                    //clicks on the search field 
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).Click();
                    //enters "c" into the field
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("C");
                    //press backspace to remove the above entered C
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys(Keys.Backspace);

                    //Enter "client Bill" into the search box
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("c");
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("l");

                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("i");

                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("e");
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("n");
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys("t");

                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys(" bill");
                    //presses backspace two times
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys(Keys.Backspace);

                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys(Keys.Backspace);
                    //puts system to sleep for 1 second
                    System.Threading.Thread.Sleep(1000);
                    //Presses enter to select the top search result loaded
                    gc.FindElement(By.XPath("//*[@id='text32v1']")).SendKeys(Keys.Enter);
                    //puts system to sleep for 10 seconds
                    System.Threading.Thread.Sleep(10000);

                }
                catch (Exception ex)
                {
                    //if an exception occurs during the report opening process in prism
                    //system generated an error log for daily report and oen for the email error log
                    //After that jumps to end of the process skipping all the process in between
                    //mail.Add("Client Bill Pending report opening failed !");
                    Exception e = new Exception("Client Bill Pending report opening failed !");
                    Console.WriteLine(ex);
                    Console.WriteLine("Client bill pending report opening failed");
                   // ErrorLogging(ex)
                   // ErrorLogging(e);
                    goto End;
                }
                try
                {
                    //puts system to sleep for one second
                    System.Threading.Thread.Sleep(1000);
                    //clicks the import button
                    gc.FindElement(By.XPath("//*[@id='button31v2']")).Click();
                    //put system to sleep for one second
                    System.Threading.Thread.Sleep(1000);
                }
                catch
                {

                }
                //gets the name of current tab of chrome and stores it into a string variable for later.
                String current = gc.CurrentWindowHandle;
                //checks if any pop has appeard.
                foreach (String winHandle in gc.WindowHandles)
                {
                    //switches to the found pop
                    gc.SwitchTo().Window(winHandle);
                }
                //sometimes the upload window(pop up) doesn't open
                if (gc.CurrentWindowHandle != current) //checks whether the pop up has appeared
                {
                    //if the pop has appeared
                    try
                    {
                        //sends the direct path of the importable text field
                        //to the file browser
                        gc.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath + "\\data.txt"); //put the full path of file here
                        //clicks the submit button
                        gc.FindElement(By.XPath("//*[@id='submit1']")).Click();
                        //waits for one second
                        System.Threading.Thread.Sleep(1000);
                        //clicks the okay button
                        gc.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                        //puts system to sleep for 20 seconds
                        System.Threading.Thread.Sleep(20000);
                        //switches control back to the main window of browser
                        gc.SwitchTo().Window(current);
                    }
                    catch
                    {
                        //if the pop failed to appear
                        try
                        {
                            //closes any opened browser instance
                            gc.Close();
                        }
                        catch
                        {

                        }
                        try
                        {
                            //tries to switch back to the main window if any and closes it
                            gc.SwitchTo().Window(current);
                            gc.Close();
                        }
                        catch
                        {

                        }
                        // goes to the end of the process
                        goto End;
                    }
                    try
                    {
                        //checks if there were an error during the import procedure or imported file
                        Exception s = new Exception(gc.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                        //logs that error for emailing
                        ErrorLogging1(s);

                    }
                    catch
                    {

                    }

                    try
                    {
                        //click the view data button
                        gc.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                        //put system to sleep for 2 seconds
                        System.Threading.Thread.Sleep(2000);
                        //clicks post data button
                        gc.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                        //waits for 2 seconds
                        System.Threading.Thread.Sleep(2000);
                        //switch to the pop and presses the okay button on it
                        gc.SwitchTo().Alert().Accept();
                        //switches back to the main window
                        gc.SwitchTo().Window(current);
                        //presses close button
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                    }
                    catch (Exception e)
                    {
                        //if view data button is not found
                        //logs the error over console
                        //switches to the main window
                        //clicks the close button
                        Console.WriteLine(e);
                        gc.SwitchTo().Window(current);
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                    }
                    tryagain:
                    try
                    {
                        //closes the browser instance
                        //destroyes the selenium browser variable

                        gc.Close();
                        gc.Dispose();
                        Console.WriteLine("Process Complete..!");

                    }
                    catch (Exception ex)
                    {
                        //if browser closing fails. 
                        //logs error through email
                        //puts sytem to sleep for 10 seconds
                        //goes up back to try again
                        Exception e = new Exception("Chrome closing failed failed !");
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }

                }
                else
                {
                    //if no importable data is found
                    tryagain:
                    //closes the browser instance
                    try
                    {
                        gc.Close();
                        gc.Dispose();
                        goto top;
                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Chrome closing failed failed !");
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }
                }

            }
            End:
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("System going to wait");

        } //main ends here
    }
}
