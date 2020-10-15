using System;
using System.Windows;
using Microsoft.Win32;
using System.ServiceProcess;
using System.Linq;
using System.Xml.Linq;
using System.Diagnostics;

namespace TIMSHelper
{
    public partial class MainWindow : Window
    {
        private void StartTIMSSQLButton_Clicked(object sender, RoutedEventArgs e)
        {
            RegistryKey reg = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\MSSQL$TIMS");
            RegistryKey reg1 = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\MSSQL$TIMS08");
            RegistryKey reg2 = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\MSSQL$TIMS14");

            if (reg == null)
            {
                reg = reg1;
            }

            if (reg1 == null)
            {
                reg = reg2;
            }

            ServiceStartMode startMode = (ServiceStartMode)reg.GetValue("Start");


            if (startMode == ServiceStartMode.Disabled)

            {
                MessageBox.Show("Sorry, the TIMS SQL Service is disabled on this PC! Please call TIMS Support to verify this is a server! Click OK to Exit TIMS Helper.");
            }
            else
            {
                ServiceController service = ServiceController.GetServices()
                    .FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS");
                ServiceController service1 = ServiceController.GetServices()
                    .FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS08");
                ServiceController service2 = ServiceController.GetServices()
                    .FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS14");

                if (service == null)
                {
                    service = service1;
                }

                if (service1 == null)
                {
                    service = service2;
                }

                if (service2 == null)
                {
                    Console.WriteLine("Not installed");
                    Console.Read();
                    MessageBox.Show("Sorry, the TIMS SQL Service was not found on this PC! Please locate your server and run TIMS Helper on the server! Click OK to Exit TIMS Helper.");
                    this.Close();
                }

                else
                    Console.WriteLine(service.Status);

                string MyService = "Checking TIMS SQL Service";
                Console.WriteLine(MyService);
                MessageBox.Show("Click OK to check the status of the TIMS SQL Server on this PC.");

                ServiceController timssql = ServiceController.GetServices().FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS");
                ServiceController timssql1 = ServiceController.GetServices().FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS08");
                ServiceController timssql2 = ServiceController.GetServices().FirstOrDefault(s => s.ServiceName == "MSSQL$TIMS14");


                if (timssql == null)
                {
                    timssql = timssql1;
                }

                if (timssql1 == null)
                {
                    timssql = timssql2;
                }

                Console.WriteLine("TIMS SQL Service is currently {0}",
                                  timssql.Status.ToString());


                switch (timssql.Status)
                {
                    case ServiceControllerStatus.Stopped:
                    case ServiceControllerStatus.StopPending:
                        {
                            // Start the service if the current status is "Stopped".

                            Console.WriteLine("Press any to start the TIMS SQL Service...");
                            timssql.Start();
                            Console.Read();
                            MessageBox.Show("TIMS SQL Server is not running. Click OK to start the TIMS SQL Server.");

                            timssql.Refresh();
                            Console.WriteLine("TIMS SQL Service is now {0}. Press any key to close...",
                                                timssql.Status.ToString());
                            Console.Read();
                            MessageBox.Show("Successfully started TIMS SQL Server! Please try to log into TIMS again. Click OK to Exit TIMS Helper.");
                            this.Close();
                            break;
                        }


                    case ServiceControllerStatus.Paused:
                    case ServiceControllerStatus.PausePending:
                        {
                            // The service is "Paused".

                            Console.WriteLine("Press any to stop the TIMS SQL Service...");
                            Console.Read();
                            timssql.Stop();

                            timssql.Refresh();
                            Console.WriteLine("TIMS SQL Service is now {0}. Press any key to close...",
                                               timssql.Status.ToString());
                            Console.Read();
                            MessageBox.Show("Starting TIMS SQL was not successful. Please try again!");
                            break;
                        }

                    case ServiceControllerStatus.Running:
                        {
                            // The service is "Running".

                            Console.WriteLine("Since TIMS SQL Service is already running, press any key to exit...");
                            Console.Read();
                            MessageBox.Show("TIMS SQL Server is already running! Sorry, please call TIMS support or try to reboot this PC or Server. Click OK to Exit TIMS Helper.");
                            break;
                        }

                }

            }
        }

        private void TIMSSlowButton_Clicked(object sender, RoutedEventArgs e)
        {
                RegistryKey Spelling =
                Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Spelling");
            string keyName = @"HKEY_CURRENT_USER\Software\Microsoft\Spelling\Dictionaries";
            string valueName = "_Global_";

            if (Registry.GetValue(keyName, valueName, null) == null)
            {
                //code if key Not Exist 
                MessageBox.Show("Global Entry is already deleted or does not exist! Sorry, no improvement!");
            }

            else
            {
                //code if key Exist
                RegistryKey badEntry = Spelling.OpenSubKey("Dictionaries", true);
                // Delete the Global Entry.
                badEntry.DeleteValue(valueName);
            }

        }

        private void NOAHErrorButton_Clicked(object sender, RoutedEventArgs e)
        {
            string MyService = "Checking NOAH Server Service";
            Console.WriteLine(MyService);

            ServiceController timssql = new ServiceController("NoahServer");
            Console.WriteLine("NOAH Server is currently {0}",
                              timssql.Status.ToString());

            switch (timssql.Status)
            {
                case ServiceControllerStatus.Stopped:
                case ServiceControllerStatus.StopPending:
                    {
                        // Start the service if the current status is "Stopped".

                        Console.WriteLine("Press any to start the TIMS SQL Service...");
                        timssql.Start();
                        Console.Read();

                        timssql.Refresh();
                        Console.WriteLine("TIMS SQL Service is now {0}. Press any key to close...",
                                            timssql.Status.ToString());
                        Console.Read();
                        MessageBox.Show("Successfully started TIMS SQL! Please try to log into TIMS again!");
                        break;
                    }


                case ServiceControllerStatus.Paused:
                case ServiceControllerStatus.PausePending:
                    {
                        // The service is "Paused".

                        timssql.Stop();

                        timssql.Refresh();
                        Console.WriteLine("NOAH Service is now {0}. Press any key to close...",
                                           timssql.Status.ToString());
                        Console.Read();
                        MessageBox.Show("NOAH was not successfully restarted. Please try again!");
                        break;
                    }

                case ServiceControllerStatus.Running:
                    {
                        // The service is "Running".

                        MessageBox.Show("NOAH Service is already running! Sorry, please call TIMS support or try to reboot server or PC!");
                        break;
                    }
            }
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void RestartReplicationButton_Clicked(object sender, RoutedEventArgs e)
        {
            RegistryKey reg = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\services\TIMSTaskManager");

            ServiceStartMode startMode = (ServiceStartMode)reg.GetValue("Start");


            if (startMode == ServiceStartMode.Disabled)

            {
                MessageBox.Show("Sorry, the TIMS Task Manager and Replication Service is disabled on this PC! Please call TIMS Support to verify this is a server!");
            }
            else
            {
                ServiceController replication = ServiceController.GetServices()
                    .FirstOrDefault(s => s.ServiceName == "TIMSTaskManager");

                if (replication == null)
                {
                    Console.WriteLine("Not installed");
                    Console.Read();
                    MessageBox.Show("Sorry, the TIMS Task Manager or Replication Service was not found on this PC! Please locate your server and run TIMS Helper on the server!");
                }
                else
                    Console.WriteLine(replication.Status);

                string MyService = "Checking TIMS Task Manager Service";
                Console.WriteLine(MyService);

                //ServiceController replication = new ServiceController("TIMSTaskManager");
                //Console.WriteLine("TIMS Task Manager Service is currently {0}",
                //                  replication.Status.ToString());


                switch (replication.Status)
                {


                    case ServiceControllerStatus.Stopped:
                    case ServiceControllerStatus.StopPending:
                        {
                            // Start the service if the current status is "Stopped".

                            Console.WriteLine("Press any to start the TIMS Task Manager Service...");
                            replication.Start();
                            Console.Read();

                            replication.Refresh();
                            Console.WriteLine("TIMS Task Manager Service is now {0}. Press any key to close...",
                                                replication.Status.ToString());
                            Console.Read();
                            MessageBox.Show("Successfully started TIMS Task Manager and Replication Service! Please try to replicate again or TIMS will try to replicate again in 15 minutes!");
                            break;
                        }


                    case ServiceControllerStatus.Paused:
                    case ServiceControllerStatus.PausePending:
                        {
                            // The service is "Paused".

                            Console.WriteLine("Press any to stop the TIMS Task Manager Service...");
                            Console.Read();
                            replication.Stop();

                            replication.Refresh();
                            Console.WriteLine("TIMS Task Manager Service is now {0}. Press any key to close...",
                                               replication.Status.ToString());
                            Console.Read();
                            MessageBox.Show("Starting TIMS Task Manager and Replication Service was not successful. Please try again!");
                            break;
                        }

                    case ServiceControllerStatus.Running:
                        {
                            // The service is "Running".

                            Console.WriteLine("Since TIMS Task Manager Service is already running, press any key to continue...");
                            Console.Read();
                            var message = "TIMS Task Manager and Replication Service is running. Do you want to restart?";
                            var title = "Restart Replication";
                            var result = MessageBox.Show(message, title, MessageBoxButton.YesNo);

                            switch (result)

                            {
                                case MessageBoxResult.Yes:   // Yes button pressed
                                    Console.WriteLine("Press any key to stop the TIMS Task Manager Service...");
                                    replication.Stop();

                                    replication.WaitForStatus(ServiceControllerStatus.Stopped);

                                    Console.WriteLine("TIMS Task Manager Service is now {0}. Press any key to continue...",
                                                      replication.Status.ToString());
                                    replication.Start();
                                    replication.WaitForStatus(ServiceControllerStatus.Running);

                                    Console.WriteLine("TIMS Task Manager Service is now {0}. Press any key to continue...",
                                                      replication.Status.ToString());
                                    MessageBox.Show("Successfully restarted TIMS Task Manager and Replication Service!");
                                    break;

                                case MessageBoxResult.No:    // No button pressed
                                    MessageBox.Show("TIMS Task Manager and Replication not restarted.");
                                    break;
                                default:                 // Neither Yes nor No pressed (just in case)
                                    MessageBox.Show("What did you press?");
                                    break;
                            }

                            break;
                        }

                }

            }
        }

        private void NOAHModulesMissingButton_Clicked(object sender, RoutedEventArgs e)

        {
            string text = System.IO.File.ReadAllText(@"C:\ProgramData\TIMSAudiology\TIMSSettings.cfg");
            System.Console.WriteLine("Contents of TIMSSettings.cfg = {0}", text);
            {
                // Example file contents
                //string texta =
                //    "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                //    "    <TIMSSettingsList><TIMSSetting Key=\"IsNoahEnabled\" Value=\"False\" />" +
                //    "</TIMSSettingsList>";

                // parse the contents into document object
                var document = XDocument.Parse(text);

                // retrieve the root object (TIMSSetttingsList)
                var settingsList = document.Root;
                // make sure it's there
                if (settingsList != null)
                {
                    // iterate through the elements inside TIMSSettingsList
                    foreach (var setting in settingsList.Elements())
                    {
                        // Check the "Key" attribute for "IsNoahEnabled" value
                        var key = setting.Attribute("Key");

                        // Is this the right setting?
                        if (key != null && key.Value == "IsNoahEnabled")
                        {
                            //Set the "Value" attribute value to true
                            var value = setting.Attribute("Value");
                            if (value != null)
                            {
                                value.Value = "True";
                            }
                        }
                    }
                }

                // Store the modified document

                text = document.ToString();
                System.Console.WriteLine("Contents of TIMSSettings.cfg = {0}", text);
                document.Save(@"C:\ProgramData\TIMSAudiology\TIMSSettings.cfg");
                MessageBox.Show("We have verified NOAH is running within TIMS now! Please log back into TIMS and verify NOAH Icons and Modules are listed. If icons are still not listed, please call TIMS support or try to reboot server or PC!");
                
            }

        }

        private void Button_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {

        }
    }

}
