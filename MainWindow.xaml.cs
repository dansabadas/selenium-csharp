using System;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
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
        NotifyIcon _notifyIcon = new NotifyIcon();

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
            _webDriver = CreateChromeDriver(); // CreatePhantomJSDriver(); // CreateChromeDriver();

            Closing += (s, e) =>
            {
                _webDriver.Dispose();
                Process.GetProcesses().Where(p => p.ProcessName == "phantomjs" || p.ProcessName == "chromedriver").ToList().ForEach(process => process.Kill());
            };

            Loaded += (s, e) =>
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                ReadPunchTimes(punchPage);

                _timer = new Timer { Interval = 1 * 60 * 1000, Enabled = true, AutoReset = true }; // every half-hour fires!
                _timer.Elapsed += (s1, e1) => VerifyTimeAgainstPunchline();
            };
        }

        private void InitializeNotificationArea()
        {
            ShowInTaskbar = false;
            WindowState = WindowState.Minimized;

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

        private void VerifyTimeAgainstPunchline()
        {
            var now = DateTime.Now;
            if (now.Hour < 8)
            {
                return;
            }

            if (now.Date > _applicationPunchState.In.Date)
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                ((IWebElement)punchPage.btnPunchIn).Click();
                return;
            }

            if (now.Subtract(_applicationPunchState.In).TotalMinutes > int.Parse(ConfigurationManager.AppSettings["minutesSpentAtOffice"]) &&   // more than 510' = 8hr30min passed and the user didn't play randomly with the online-punch
                _applicationPunchState.In > _applicationPunchState.Out)    
            {
                var punchPage = NavigateToPunchPageAndCollectUIControls();
                ((IWebElement)punchPage.btnPunchOut).Click();
            }
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
            ReadPunchTimes(punchPage);
        }

        private void ReadPunchTimes(dynamic punchPage)
        {
            _applicationPunchState.In = DateTime.ParseExact(((IWebElement)punchPage.lblStart).Text.Trim(), "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture);
            _applicationPunchState.Out = DateTime.ParseExact(((IWebElement)punchPage.lblEnd).Text.Trim(), "dd.MM.yyyy HH:mm", CultureInfo.InvariantCulture);
        }

        private IWebDriver CreateInternetExplorerDriver()
        {
            var ieOptions = new InternetExplorerOptions
            {
                IgnoreZoomLevel = true,
                IntroduceInstabilityByIgnoringProtectedModeSettings = true
            };

            return new InternetExplorerDriver(ieOptions);
        }

        private static IWebDriver CreateChromeDriver()
        {
            return new ChromeDriver();
        }

        private static IWebDriver CreatePhantomJSDriver()
        {
            return new PhantomJSDriver();
        }

        private void FillCredentials()
        {
            _webDriver.FindElement(By.Id("txtUser")).SendKeys(ConfigurationManager.AppSettings["userName"]);
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
    }
}
