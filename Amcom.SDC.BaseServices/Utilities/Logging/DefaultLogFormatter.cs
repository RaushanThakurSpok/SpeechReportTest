using System;
using System.Collections.Generic;
using System.Text;

namespace Utilities.Logging
{
    /// <summary>
    /// 
    /// </summary>
    public class DefaultLogFormatter : LogFormatter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultLogFormatter"/> class.
        /// </summary>
        public DefaultLogFormatter()
        {
            FormatString = "[{0}{4} {1,-7}{3}] {2}";
        }

        /// <summary>
        /// Does the format entry.
        /// </summary>
        /// <param name="severity">The severity.</param>
        /// <param name="timestamp">The timestamp.</param>
        /// <param name="utcOffset">The utc offset.</param>
        /// <param name="logEntry">The log entry.</param>
        /// <param name="category">The category.</param>
        /// <param name="data">The data.</param>
        /// <returns></returns>
        protected override string DoFormatEntry(LoggerSeverity severity, DateTime timestamp, TimeSpan? utcOffset, string logEntry, string category, Dictionary<string, object> data)
        {
            if (!string.IsNullOrEmpty(category))
            {
                category = " " + category;
            }
            return string.Format(FormatString, timestamp, severity, logEntry, category, DateTimeUtlities.TrimTimeSpan(DateTimeUtlities.ToStringTimeSpan(utcOffset)));
        }
    }
}
