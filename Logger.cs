using System;
using log4net;

namespace Logger
{
    public class Logger
    {

        public enum LogType
        {
            Info = 1,
            Debug = 2,
            Fatal = 3
        }


        private static ILog _logger;
        private static Logger _instance;
        private Logger() { }
        public static Logger Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new Logger();
                }
                return _instance;
            }
        }

        public static void SetLogger(string path, string programType)
        {
            if (_logger == null)
            {
                _logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                log4net.GlobalContext.Properties["LogFileName"] = path + "YouPrice." + programType + "_" + System.Environment.UserName + "_" + System.Environment.MachineName + "_" + DateTime.Now.ToString("yyyyMMdd_Hmmss") + ".log";
                log4net.Config.XmlConfigurator.Configure();
            }

        }

        public void Log(string message, LogType logType = LogType.Info)
        {
            string timePrefix = DateTime.Now.ToString("H:mm:ss") + " - ";
            switch (logType)
            {
                case LogType.Debug:
                    _logger.Debug(message);
                    Console.WriteLine(timePrefix + "Debug : " + message);
                    break;
                case LogType.Fatal:
                    _logger.Fatal(message);
                    Console.WriteLine(timePrefix + "Fatal : " + message);
                    break;
                default:
                    _logger.Info(message);
                    Console.WriteLine(timePrefix + "Info : " + message);
                    break;
            }

            
        }
    }
}
