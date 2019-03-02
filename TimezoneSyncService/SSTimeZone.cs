using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Xml;
using System.Linq;


namespace SSTimezone
{
    public partial class SSTimeZone : ServiceBase
    {

        string sEventLog;                   // eventlog (APplication) object for the service for writing log entries.

        System.Timers.Timer tzTimer;        // timer object for the service.
        
        // Registry Settings:
        string sXMLPath;                    // path read from registry pointing to the xml schedule file
        string sFrequencyMinutes;           // integer read from registry
        Double dblInterval;                 // milliseconds derived from the registry/minutes setting
        bool bIsGMT;                        // registry setting for IsGMT
        bool bVerboseLogging;               // registry setting for VerboseLogging
        string sLastUpdatedStartDatetime;   // the start_datetime value from the xml file when the last change occurred;
        
        // service Globals
        string sTimeZoneNew;                // will contain the new timezone to change to if necessary
        System.Data.DataSet ds;             // Dataset that contains the contents of the xml file
        string s64Bit = "";                 // registry string in the event of 64 bit


        public SSTimeZone()
        {
            InitializeComponent();
            this.tzTimer = new System.Timers.Timer(20000);
            this.tzTimer.Elapsed += new System.Timers.ElapsedEventHandler(tzTimer_Elapsed);
            this.tzTimer.Enabled = true;
            this.tzTimer.AutoReset = true;
        }

        void tzTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (this.bIsGMT)
            {
                // Automatically set the timezone to UTC Casablanca (Morocco Standard Time) if not already
                this.sTimeZoneNew = "Morocco Standard Time";
                if(System.TimeZoneInfo.Local.StandardName!=this.sTimeZoneNew)
                    SetLocalTimezone();
            }       
            else
            {
                //this.tzEventLog.WriteEntry("timer");
                this.sTimeZoneNew = "";
                CheckUpdateTimezone();
            }
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            this.GetRegistrySettings();
            this.tzTimer.Interval = dblInterval;
            this.tzTimer.Start();
            this.tzEventLog.WriteEntry("Timezone Sync Service Version " + System.Windows.Forms.Application.ProductVersion + " Startup" +
                "\n\n Registry settings:" +
                "\n Frequency:         " + this.sFrequencyMinutes.ToString() + "min / " + dblInterval.ToString() + "ms." +
                "\n XMLPath:           " + sXMLPath +
                "\n IsGMT:             " + this.bIsGMT.ToString() +
                "\n VerboseLogging:    " + this.bVerboseLogging.ToString());

            this.Is64Bit();
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            this.tzTimer.Stop();
            this.tzTimer.Dispose();
            this.tzEventLog.WriteEntry("Timezone Sync Service Version " + System.Windows.Forms.Application.ProductVersion + " Stopped.");
        }

        private bool Is64Bit()
        {
            bool bRetval = false;

            if ((Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE") == "AMD64") || (Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432") == "AMD64"))
            {
                bRetval = true;
                s64Bit = @"Wow6432Node\";
            }
            return bRetval;
        }

        private void GetRegistrySettings()
        {
            try
            {
                //XMLPath
                if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "XMLPath", "PathNotFound"))
                {
                    this.sXMLPath = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "XMLPath", "PathNotFound").ToString();
                }
                else if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "XMLPath", "PathNotFound"))
                {
                    this.sXMLPath = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "XMLPath", "PathNotFound").ToString();
                }
                else
                {
                    this.sXMLPath = @"XML Path NOT FOUND. Please check registry settings.";
                }

                //FrequencyMinutes
                if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "FrequencyMinutes", "-1"))
                {
                    this.sFrequencyMinutes = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "FrequencyMinutes", "-1").ToString();
                }
                else if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "FrequencyMinutes", "-1"))
                {
                    this.sFrequencyMinutes = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "FrequencyMinutes", "-1").ToString();
                }
                else
                {
                    this.sFrequencyMinutes = @"-1";
                }

                //IsGMT
                if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "IsGMT", "0"))
                {
                    this.bIsGMT = "1" == Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "IsGMT", "0").ToString();
                }
                else if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "IsGMT", "0"))
                {
                    this.bIsGMT = "1" == Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "IsGMT", "0").ToString();
                }
                else
                {
                    this.bIsGMT = false;
                }

                //VerboseLogging
                if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "VerboseLogging", "1"))
                {
                    this.bVerboseLogging = "1" == Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "VerboseLogging", "1").ToString();
                }
                else if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "VerboseLogging", "1"))
                {
                    this.bVerboseLogging = "1" == Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "VerboseLogging", "1").ToString();
                }
                else
                {
                    this.bVerboseLogging = true;
                }

                //LastUpdatedStartDatetime
                if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "LastUpdatedStartDatetime", "-1"))
                {
                    this.sLastUpdatedStartDatetime = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\SilverSea\Timezone Sync Service", "LastUpdatedStartDatetime", "-1").ToString();
                }
                else if (null != Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "LastUpdatedStartDatetime", "-1"))
                {
                    this.sLastUpdatedStartDatetime = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\SilverSea\Timezone Sync Service", "LastUpdatedStartDatetime", "-1").ToString();
                }
                else
                {
                    this.sLastUpdatedStartDatetime = @"-1";
                }

                
                if (Double.TryParse(sFrequencyMinutes, out dblInterval))
                {
                    dblInterval = dblInterval * 60000;
                }
                else
                {
                    dblInterval = 600000;
                }
            }
            catch (Exception xe)
            {
                this.sXMLPath = xe.Message;
                sFrequencyMinutes = "-1";
                dblInterval = 600000;
                this.bVerboseLogging = true;
                this.bIsGMT = false;
                this.sLastUpdatedStartDatetime = @"-1";
                this.tzEventLog.WriteEntry("Exception: GetRegistrySettings(): \n" + xe.Message);
            }
        }

        private bool CheckUpdateTimezone()
        {
            try
            {
                bool bContinue = true;
                
                sEventLog = "Timezone Sync Service Version " + System.Windows.Forms.Application.ProductVersion;
                sEventLog += "\nChecking Timezone:";
                sEventLog += "\n\tLocal Client Timezone: " + System.TimeZone.CurrentTimeZone.StandardName;
                sEventLog += "\n\tMost recent active timezone entry from xml file: " + sLastUpdatedStartDatetime;
                sEventLog += "\n\tLoading TimezoneSchedule.xml: " + sXMLPath;

                bContinue = LoadXMLSchedule();
                if (!bContinue) throw new Exception("LoadXMLSchedule error");

                DataView dv;
                DataTable dt;

                if (bContinue)
                {
                    // read first table of global params into temp datatable
                    dv = this.ds.Tables[0].DefaultView;
                    dt = dv.ToTable();

                    // check xmlfile for any new settings that need to be saved into Registry settings:
                    if (dt.Rows[0]["new_XMLPath"].ToString().Length > 0)
                    {
                        this.sXMLPath = dt.Rows[0]["new_XMLPath"].ToString();
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\" + s64Bit + @"SilverSea\Timezone Sync Service", "XMLPath", this.sXMLPath);
                    }

                    if (dt.Rows[0]["set_VerboseLogging"].ToString().Length > 0)
                    {
                        this.bVerboseLogging = ("1"==dt.Rows[0]["set_VerboseLogging"].ToString());
                        if(this.bVerboseLogging)
                            Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\" + s64Bit + @"SilverSea\Timezone Sync Service", "VerboseLogging", "1");
                        else
                            Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\" + s64Bit + @"SilverSea\Timezone Sync Service", "VerboseLogging", "0");
                    }

                    if (dt.Rows[0]["new_FrequencyMinutes"].ToString().Length > 0)
                    {
                        this.sFrequencyMinutes = dt.Rows[0]["new_FrequencyMinutes"].ToString();
                        Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\" + s64Bit + @"SilverSea\Timezone Sync Service", "FrequencyMinutes", this.sFrequencyMinutes);

                        // if the frequency changed, parse that into milliseconds and restart the timer with the new value:
                        if (Double.TryParse(sFrequencyMinutes, out dblInterval))
                        {
                            dblInterval = dblInterval * 60000;
                        }
                        else
                        {
                            dblInterval = 600000;
                        }
                        this.tzTimer.Stop();
                        this.tzTimer.Interval = dblInterval;
                        this.tzTimer.Start();
                    }
                }

                //reload registry settings to obtain the LastUpdatedStartDatetime from registry and make the new changes go into effect now
                this.GetRegistrySettings();

                // check the LastUpdatedStartDatetime to see if it is in the future, if so then we went hour back and should not check again.
                // the system time has to be ahead of the LastUpdatedStartDatetime in order to check!
                if ("-1" != sLastUpdatedStartDatetime)
                {
                    bContinue = DateTime.Compare(DateTime.Now, DateTime.Parse(sLastUpdatedStartDatetime)) >= 0;
                    if (!bContinue) sEventLog += "\n\nSystem datetime is earlier than the last recorded actual time zone change; deferring until next check";
                }
                
                //Read the file and call the correct time zone change required check function
                if (bContinue)
                {

                    // read first table of global params into temp datatable
                    dv = this.ds.Tables[0].DefaultView;
                    dt = dv.ToTable(); 
                    
                    if ("true" == dt.Rows[0]["IsRecurring"].ToString())
                    {
                        // 7-day repeating itinerary, using <day_of_week> and <start_time> in xml file:
                        bContinue = this.TimezoneChangeRequired_Recurring();
                        sEventLog += "\n\nTimezone change required (Recurring Schedule): \n\t" + bContinue.ToString();
                    }
                    else
                    {
                        // Ongoing non-recurring itinerary, using <start_datetime> in xml file:
                        bContinue = this.TimezoneChangeRequired_Indefinite();
                        sEventLog += "\n\nTimezone change required (Indefinite Schedule): \n\t" + bContinue.ToString();
                    }
                    
                }

                if (bContinue)
                {
                    SetLocalTimezone();
                }

                // for verbose logging, write the long entry:
                if (this.bVerboseLogging || bContinue)
                {
                    this.tzEventLog.WriteEntry(sEventLog);
                }

                return bContinue;

            }
            catch (Exception tzExc)
            {
                this.tzEventLog.WriteEntry("Exception: CheckUpdateTimezone(): \nEvent Log:\n" + sEventLog + "\n\nError:\n" + tzExc.Message);
                return false;
            }        
        }

        private bool LoadXMLSchedule()
        {
            try
            {   
                this.ds = new DataSet("Timezones");
                this.ds.ReadXml(this.sXMLPath);

                sEventLog += "\n\tLoadXML Tables count: " + this.ds.Tables.Count.ToString();
                sEventLog += "\n\tLoadXML Rows count <entry>: " + this.ds.Tables[1].Rows.Count.ToString();
                
                return true;
            }
            catch (XmlException xe)
            {
                this.tzEventLog.WriteEntry("Exception LoadXMLSchedule(): \n" + xe.Message + "\n" + xe.StackTrace);
                return false;
            }
        }

        private bool TimezoneChangeRequired_Recurring() 
        {
            try
            {

                int nSavedIndex = 0;
                int nCurrentDoW = 0;
                TimeSpan tsCurrentHHMM;
               
                // sort array by dow, start_time ASC
                DataView dv = this.ds.Tables[1].DefaultView;
                dv.Sort = "day_of_week, start_time ASC";

                DataTable dt = dv.ToTable();

                // loop through the data set
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    nCurrentDoW = int.Parse(dt.Rows[i]["day_of_week"].ToString());

                    // if the dow in the first row is < the actual dow, save the index (overwrite previous).
                    if (nCurrentDoW < (int)System.DateTime.Now.DayOfWeek)
                    {
                        nSavedIndex = i;
                    }
                    // if the dow = the actual dow, check start_time.
                    else if (nCurrentDoW == (int)System.DateTime.Now.DayOfWeek)
                    {
                        tsCurrentHHMM = System.TimeSpan.Parse(dt.Rows[i]["start_time"].ToString());

                        // if start_time is <= the actual hhmm, save the index (overwrite previous).
                        if (tsCurrentHHMM.Hours <= System.DateTime.Now.TimeOfDay.Hours)
                        {
                            nSavedIndex = i;
                        }
                        else
                        {
                            // if start_time is > actual hhmm, do nothing, break out of loop
                            break;
                        }
                    }
                    else
                    {
                        // if the dow is greater than the actual dow, do nothing, break out of loop.
                        break;
                    }

                }


                // once out of loop,
                // the saved index is the active scheduled time zone we should be on.
                // obtain the standard name from the active scheduled time zone
                this.sTimeZoneNew = dt.Rows[nSavedIndex]["tzStandardName"].ToString();
                this.sLastUpdatedStartDatetime = dt.Rows[nSavedIndex]["start_datetime"].ToString();

                sEventLog += "\n\nActive Timezone Entry from XML File: \n\t" + DateTime.Parse(dt.Rows[nSavedIndex]["start_datetime"].ToString()).ToString() + " / " + dt.Rows[nSavedIndex]["tzStandardName"].ToString();

                sEventLog += "\n\nSystem.TimeZone.CurrentTimeZone.StandardName: \n\t" + System.TimeZone.CurrentTimeZone.StandardName;

                //compare to actual time zone and act accordingly.

                return String.Compare(this.sTimeZoneNew, System.TimeZone.CurrentTimeZone.StandardName) == 0;


            }
            catch (Exception excTCR)
            {
                this.tzEventLog.WriteEntry("Exception TimezoneChangeRequired_Recurring(): \n" + excTCR.Message + "\n" + excTCR.StackTrace);
                return false;
            }


        }

        private bool TimezoneChangeRequired_Indefinite()
        {
            try
            {

                int nSavedIndex = 0;
                DateTime dtmCurrentDate;

                string strTZActive = "";

                // sort array by startdate, start_time ASC
                DataView dv = this.ds.Tables[1].DefaultView;
                dv.Sort = "start_datetime ASC";

                DataTable dt = dv.ToTable();

                // loop through the data set
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dtmCurrentDate = DateTime.Parse(dt.Rows[i]["start_datetime"].ToString());

                    // if the dow in the first row is < the actual dow, save the index (overwrite previous).
                    if (dtmCurrentDate <= System.DateTime.Now)
                    {
                        nSavedIndex = i;
                    }
                    else
                    {
                        // if the startdate_YYYYMMDD is greater than the actual system start date, do nothing, break out of loop.
                        break;
                    }

                }


                // once out of loop,
                // the saved index is the active scheduled time zone we should be on.
                // obtain the standard name from the active scheduled time zone
                strTZActive = dt.Rows[nSavedIndex]["tzStandardName"].ToString();
                this.sLastUpdatedStartDatetime = dt.Rows[nSavedIndex]["start_datetime"].ToString();

                sEventLog += "\n\nActive Timezone Entry from XML File: \n\t" + DateTime.Parse(dt.Rows[nSavedIndex]["start_datetime"].ToString()).ToString() + " / " + dt.Rows[nSavedIndex]["tzStandardName"].ToString();

                //compare to actual time zone and act accordingly.
                sEventLog += "\n\nSystem.TimeZone.CurrentTimeZone.StandardName: \n\t" + System.TimeZone.CurrentTimeZone.StandardName;

                if (strTZActive != System.TimeZone.CurrentTimeZone.StandardName)
                {
                    this.sTimeZoneNew = strTZActive;
                    return true;
                }
                else
                {
                   return false;
                }

            }
            catch (Exception excTCR)
            {
                this.tzEventLog.WriteEntry("Exception TimezoneChangeRequired_Indefinite():\n" + excTCR.Message + "\n" + excTCR.StackTrace);
                return false;
            }
        }


        private void SetLocalTimezone_PInvoke()
        {
          
            try 
            {
                // * USING INTEROP
                //create timezone struct info for new tzi
                SSTimezone.TimezoneFunctionality.TimeZoneInformation tzi = new TimezoneFunctionality.TimeZoneInformation();
                
                SSTimezone.TimezoneFunctionality.SystemTime tziDtmStandard = new TimezoneFunctionality.SystemTime();
                SSTimezone.TimezoneFunctionality.SystemTime tziDtmDaylight = new TimezoneFunctionality.SystemTime();

                
                //get the new timezoneinfo for the new timezone
                System.TimeZoneInfo tziNew = System.TimeZoneInfo.FindSystemTimeZoneById(this.sTimeZoneNew);

             
                TimeZoneInfo.AdjustmentRule[] adjustmentRules = tziNew.GetAdjustmentRules();
                TimeZoneInfo.AdjustmentRule adjustmentRule = null;
                if (adjustmentRules.Length > 0)
                {
                    // Find the single record that encompasses today's date. If none exists, sets adjustmentRule to null.
                    adjustmentRule = adjustmentRules.SingleOrDefault(ar => ar.DateStart <= DateTime.Now && DateTime.Now <= ar.DateEnd);
                }


                double bias = -tziNew.BaseUtcOffset.TotalMinutes; // I'm not sure why this number needs to be negated, but it does.
                double daylightBias = adjustmentRule == null ? -60 : -adjustmentRule.DaylightDelta.TotalMinutes; // Not sure why default is -60, or why this number needs to be negated, but it does.
               
                
                int daylightDay = 0;
                int daylightDayOfWeek = 0;
                int daylightHour = 0;
                int daylightMonth = 0;
                int standardDay = 0;
                int standardDayOfWeek = 0;
                int standardHour = 0;
                int standardMonth = 0;

                if (adjustmentRule != null)
                {
                    TimeZoneInfo.TransitionTime daylightTime = adjustmentRule.DaylightTransitionStart;
                    TimeZoneInfo.TransitionTime standardTime = adjustmentRule.DaylightTransitionEnd;

                    // Valid values depend on IsFixedDateRule: http://msdn.microsoft.com/en-us/library/system.timezoneinfo.transitiontime.isfixeddaterule.
                    daylightDay = daylightTime.IsFixedDateRule ? daylightTime.Day : daylightTime.Week;
                    daylightDayOfWeek = daylightTime.IsFixedDateRule ? -1 : (int)daylightTime.DayOfWeek;
                    daylightHour = daylightTime.TimeOfDay.Hour;
                    daylightMonth = daylightTime.Month;

                    standardDay = standardTime.IsFixedDateRule ? standardTime.Day : standardTime.Week;
                    standardDayOfWeek = standardTime.IsFixedDateRule ? -1 : (int)standardTime.DayOfWeek;
                    standardHour = standardTime.TimeOfDay.Hour;
                    standardMonth = standardTime.Month;
                }
               
                

                tzi.bias = (int)bias;
                
                tzi.standardName = tziNew.StandardName;

                tziDtmStandard.day = (short)standardDay;             
                tziDtmStandard.dayOfWeek = (short)standardDayOfWeek;  
                tziDtmStandard.hour = (short)standardHour;  
                tziDtmStandard.milliseconds = 0;       
                tziDtmStandard.minute = 0;                          
                tziDtmStandard.month = (short)standardMonth;   
                tziDtmStandard.second = 0;                     

                tzi.standardDate = tziDtmStandard;

                tzi.daylightBias = (int)daylightBias;

                sEventLog += "\n\ncurrentRule.DaylightDelta.Minutes:  " + tzi.daylightBias.ToString();

                tziDtmDaylight.day = (short)daylightDay;      
                tziDtmDaylight.dayOfWeek = (short)daylightDayOfWeek;          
                tziDtmDaylight.hour = (short)daylightHour;                    
                tziDtmDaylight.milliseconds = 0;                              
                tziDtmDaylight.minute = 0;                                    
                tziDtmDaylight.month = (short)daylightMonth;                  
                tziDtmDaylight.second = 0;                                    

                tzi.daylightDate = tziDtmDaylight;

                tzi.daylightName = tziNew.DaylightName;


                //Set the local Timezone with the new information - 
                
                //check to see if this is Windows XP/2003 or Windows 7/2008, call correct time change function
                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    if (Environment.OSVersion.Version.Major == 5)
                    {
                        sEventLog += "\n\tCall to TimezoneFunctionality.SetTimeZone(tzi)";
                        TimezoneFunctionality.SetTimeZone(tzi);
                    }

                    if (Environment.OSVersion.Version.Major > 5)
                    {
                        sEventLog += "\n\tCall toTimezoneFunctionality.SetDynamicTimeZone(tzi)";
                        TimezoneFunctionality.SetDynamicTimeZone(tzi);
                    }
                }


                int nErr = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                sEventLog += "\n\nLast Win32 Error (if any): " + nErr.ToString();
                

            }
            catch (Exception tz)
            {
                sEventLog += "Exception: SetLocalTimezone_PInvoke(): " + tz.Message;
            }
        }

        private void SetLocalTimezone_WinXP2003()
        {
            try
            {
                
                //RUNDLL32.EXE SHELL32.DLL,Control_RunDLL TIMEDATE.CPL,,/Z Pacific Standard Time
                var proc = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "RUNDLL32.EXE",
                        Arguments = "SHELL32.DLL,Control_RunDLL TIMEDATE.CPL,,/Z " + this.sTimeZoneNew,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };

                proc.Start();
                while (!proc.StandardOutput.EndOfStream)
                {
                    sEventLog += "\n\n" + proc.StandardOutput.ReadLine();
                }

            }
            catch (Exception tz)
            {
                sEventLog += "Exception: SetLocalTimezone_WinXP2003(): " + tz.Message;
            }
        }

        private void SetLocalTimezone_Win7() 
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("tzutil", "/s \"" + sTimeZoneNew + "\"");
                psi.RedirectStandardOutput = true;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;

                Process pr = new Process();
                pr.StartInfo = psi;
                pr.Start();

                sEventLog += "\n\n" + pr.StandardOutput.ReadToEnd() + "\n\n";

                pr.Close();
            }
            catch (Exception tz)
            {
                sEventLog += "Exception: SetLocalTimezone_Win72008(): " + tz.Message;
            }
        
        }

            
        private void SetLocalTimezone() 
        {            
            try
            {
                sEventLog += "\n\nChanging time zone from: \n\t" +
                    System.TimeZone.CurrentTimeZone.StandardName +
                    "\n\nTo new time zone: \n\t" + sTimeZoneNew;    

                sEventLog += "\n\nOS Version: " + Environment.OSVersion.ToString();

                ////check to see if this is Windows XP/2003 or Windows 7/2008, call correct time change function
                //if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                //{
                //    if (Environment.OSVersion.Version.Major == 5)
                //    {
                //        if (Environment.OSVersion.Version.Minor > 0)
                //        {
                //            sEventLog += "\n\tCall to SetLocalTimezone_WinXP_2003()";
                //            SetLocalTimezone_WinXP2003();
                //        }
                //    }

                // Windows 7 x86

                

                if (OSInfo.Name == "Windows 7")
                {
                        // 6.1 = Windows 7
                        sEventLog += "\n\tCall to SetLocalTimezone_Win7()";
                        SetLocalTimezone_Win7();
                }
                else
                {
                        // 6.0 = Vista or Windows 2008  
                        sEventLog += "\n\tCall to SetLocalTimezone_PInvoke()";
                        SetLocalTimezone_PInvoke();
                }
                
                System.Globalization.CultureInfo.CurrentCulture.ClearCachedData();

                //update the LastUpdatedStartDatetime timestamp in the registry:
                Microsoft.Win32.Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\" + s64Bit + @"SilverSea\Timezone Sync Service", "LastUpdatedStartDatetime", this.sLastUpdatedStartDatetime);
            
            }
            catch (Exception tze)
            {
                sEventLog += "Exception: SetLocalTimezone(): " + tze.Message;
            }
        }
    }
}
