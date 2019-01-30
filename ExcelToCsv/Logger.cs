using NLog;

namespace ExcelToCsv
{
    public class Logger
    {
        private static Logger _instance;
        private readonly NLog.Logger _logger;

        private Logger()
        {
            _logger = LogManager.GetCurrentClassLogger();
        }

        private static Logger Instance => _instance ?? (_instance = new Logger());

        public static NLog.Logger Log => Instance._logger;
    }
}