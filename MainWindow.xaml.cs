using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using Timer = System.Timers.Timer;

namespace selenium_csharp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static readonly object __syncLock = new object();

        readonly NotifyIcon _notifyIcon = new NotifyIcon();

        private ApplicationPunchState _applicationPunchState;
        private IWebDriver _webDriver;
        private Timer _timer;

        public MainWindow()
        {
            InitializeComponent();

            Initialize();
        }

        private void Initialize()
        {
            InitializeNotificationArea();

            _applicationPunchState = new ApplicationPunchState();

            InitializeUIControlsDealingWithApplicationState();

            Closing += (s, e) =>
            {
                _webDriver?.Dispose();
                Process.GetProcesses().Where(p => p.ProcessName == "phantomjs" || p.ProcessName == "chromedriver").ToList().ForEach(process => process.Kill());
            };

            Loaded += (s, e) =>
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                LoadPunchTimesIntoApplicationState((IWebElement)punchPage.lblStart, (IWebElement)punchPage.lblEnd);

                _timer = new Timer { Interval = _applicationPunchState.PollingIntervalMinutes * 60 * 1000, Enabled = true, AutoReset = true }; // fires every half-hour (or whatever time interval we set in UI)!
                _timer.Elapsed += (s1, e1) =>
                {
                    lock (__syncLock)
                    {
                        PunchAndRefreshApplicationState();
                    }

                    BindApplicationStateToUI();
                };
            };

            SystemEvents.SessionSwitch += (s, e) =>
            {
                if (e.Reason == SessionSwitchReason.SessionUnlock)  // this is useful first thing in the morning at computer login (or better said unlock)
                {
                    lock (__syncLock)
                    {
                        PunchAndRefreshApplicationState();
                    }

                    BindApplicationStateToUI();
               }
            };
        }

        private void PunchAndRefreshApplicationState()
        {
            ApplyPunchingLogic();
            var punchPage = NavigateToPunchPageAndCollectUIControls();
            LoadPunchTimesIntoApplicationState((IWebElement)punchPage.lblStart, (IWebElement)punchPage.lblEnd);
        }

        private void InitializeUIControlsDealingWithApplicationState()
        {
            textBlockStartWorkingHour.LostFocus += HandleTextBox_TextInput;
            textBlockPollingIntervalMinutes.LostFocus += HandleTextBox_TextInput;
            textBlockOfficeTimeMinutes.LostFocus += HandleTextBox_TextInput;
            PopulateComboBox();
            comboBox.SelectionChanged += ComboBox_SelectionChanged;

            LoadApplicationStateSettingsFromUI();
            LoadCorrespondingWebDriver();
        }

        private void PopulateComboBox()
        {
            var data = new List<string> {"Chrome", "PhantomJS"};
            comboBox.ItemsSource = data;
            comboBox.SelectedIndex = 1;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadCorrespondingWebDriver();
        }

        private void LoadCorrespondingWebDriver()
        {
            _webDriver?.Dispose();
            switch (comboBox.SelectedItem as string)
            {
                case "Chrome":
                    _webDriver = new ChromeDriver();
                    break;
                case "PhantomJS":
                    _webDriver = new PhantomJSDriver();
                    break;
                case "InternetExplorer":
                    var ieOptions = new InternetExplorerOptions
                    {
                        IgnoreZoomLevel = true,
                        IntroduceInstabilityByIgnoringProtectedModeSettings = true
                    };

                    _webDriver = new InternetExplorerDriver(ieOptions);
                    break;
            }
        }

        private void InitializeNotificationArea()
        {
            ShowInTaskbar = false;
            WindowState = WindowState.Minimized;
            WindowStyle = WindowStyle.None;

            _notifyIcon.Visible = true;
            _notifyIcon.Icon = Properties.Resources.Custom_Icon_Design_Pretty_Office_4_ICO;
            var contextMenu = new ContextMenuStrip();
            var openMonitor = new ToolStripMenuItem();
            var exitApplication = new ToolStripMenuItem();

            _notifyIcon.ContextMenuStrip = contextMenu;

            openMonitor.Text = "Open";
            openMonitor.Click += (s, a) => WindowState = WindowState.Maximized;
            contextMenu.Items.Add(openMonitor);

            exitApplication.Text = "Exit";
            exitApplication.Click += (s, a) => Close();
            contextMenu.Items.Add(exitApplication);
        }

        private void HandleTextBox_TextInput(object sender, RoutedEventArgs e)
        {
            LoadApplicationStateSettingsFromUI();
            _timer.Interval = _applicationPunchState.PollingIntervalMinutes * 60 * 1000;    // the timer interval is expressed in milliseconds
        }

        private void LoadApplicationStateSettingsFromUI()
        {
            _applicationPunchState.PollingIntervalMinutes = int.Parse(textBlockPollingIntervalMinutes.Text);
            _applicationPunchState.StartWorkingHour = int.Parse(textBlockStartWorkingHour.Text);
            _applicationPunchState.MandatoryTimeSpentAtOfficeMinutes = int.Parse(textBlockOfficeTimeMinutes.Text);
        }

        private bool ApplyPunchingLogic()
        {
            var now = DateTime.Now;
            if (now.DayOfWeek == DayOfWeek.Saturday || now.DayOfWeek == DayOfWeek.Sunday)
            {
                return false;
            }

            if (now.Hour < _applicationPunchState.StartWorkingHour)
            {
                return false;
            }

            if (now.Date > _applicationPunchState.In.Date)
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                ((IWebElement)punchPage.btnPunchIn).Click();
                return true;
            }

            if (now.Subtract(_applicationPunchState.In).TotalMinutes > _applicationPunchState.MandatoryTimeSpentAtOfficeMinutes &&   // more than 510' = 8hr30min passed and the user didn't play randomly with the online-punch
                _applicationPunchState.In > _applicationPunchState.Out)    
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                ((IWebElement)punchPage.btnPunchOut).Click();
                return true;
            }

            return false;
        }

        private dynamic NavigateToPunchPageAndCollectUIControls()
        {
            _webDriver.Url = "https://robuhp99011v02.ad001.siemens.net/TrueHR2/Login.aspx";
            ClickElement("chkWindowsAuthentication");
            FillCredentials();
            ClickElement("btnLogin");
            var pontajLink = FindPontajLink();
            pontajLink.Click();
            var punchPage = FindPunchPageElements();

            return punchPage;
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            //IWebDriver driver = new FirefoxDriver();
            //http://selenium-release.storage.googleapis.com/index.html?path=3.6/

            var punchPage = NavigateToPunchPageAndCollectUIControls();
            LoadPunchTimesIntoApplicationState((IWebElement)punchPage.lblStart, (IWebElement)punchPage.lblEnd);
            BindApplicationStateToUI();
        }

        private void BindApplicationStateToUI()
        {
            var now = DateTime.Now;
            var minutesPassed = (int)now.Subtract(_applicationPunchState.In).TotalMinutes;
            Dispatcher.Invoke(() =>
            {
                labelStartTime.Content = _applicationPunchState.In.ToString("d.MM.yyyy HH:mm");
                labelEndTime.Content = _applicationPunchState.Out.ToString("d.MM.yyyy HH:mm");
                labelRefreshInfo.Content = $"Refreshed at {now:dd.MM.yyyy HH:mm:ss} \n{minutesPassed} minutes passed since last punch-in";
            });
        }

        private void LoadPunchTimesIntoApplicationState(IWebElement lblStart, IWebElement lblEnd)
        {
            _applicationPunchState.In = DateTime.ParseExact(lblStart.Text.Trim(), "d.MM.yyyy HH:mm", CultureInfo.InvariantCulture);
            _applicationPunchState.Out = DateTime.ParseExact(lblEnd.Text.Trim(), "d.MM.yyyy HH:mm", CultureInfo.InvariantCulture);
        }

        private void FillCredentials()
        {
            _webDriver.FindElement(By.Id("txtUser")).SendKeys(Environment.UserName);
            _webDriver.FindElement(By.Id("txtPassword")).SendKeys(ConfigurationManager.AppSettings["password"]);
        }

        private void ClickElement(string htmlElementId)
        {
            _webDriver.FindElement(By.Id(htmlElementId)).Click();
        }

        private IWebElement FindPontajLink()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(5));
            var pontajLink = wait.Until(d =>
            {
                SwitchToNamedFrame("Content");
                var element = _webDriver.FindElements(By.TagName("a")).SingleOrDefault(a => a.Text == "Pontaj");
                return element; 
            });

            return pontajLink;
        }

        private dynamic FindPunchPageElements()
        {
            IWebElement btnPunchIn = null, btnPunchOut = null, lblStart = null, lblEnd = null;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(5));
            wait.Until(d =>
            {
                SwitchToNamedFrame("WorkPage");
                btnPunchIn = _webDriver.FindElement(By.Id("btnPunchIn"));
                btnPunchOut = _webDriver.FindElement(By.Id("btnPunchOut"));
                lblStart = _webDriver.FindElement(By.Id("lblStart"));
                lblEnd = _webDriver.FindElement(By.Id("lblEnd"));
                return btnPunchIn != null;
            });

            return new { btnPunchIn, btnPunchOut, lblStart, lblEnd };
        }

        private void SwitchToNamedFrame(string frameName)
        {
            _webDriver.SwitchTo().Window(_webDriver.CurrentWindowHandle);
            var element = _webDriver.FindElement(By.Id("MainFS"));
            element = element?.FindElement(By.Id("Second"));
            element = element?.FindElement(By.Name(frameName));
            _webDriver.SwitchTo().Frame(element);
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void fastPunchButton_Click(object sender, RoutedEventArgs e)
        {
            var punchingApplied = ApplyPunchingLogic();
            if (!punchingApplied)
            {
                System.Windows.MessageBox.Show("It is not yet time to manually punch! \nMaybe change your settings!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
