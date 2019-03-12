using System;
using System.Collections.Generic;
using System.Text;

namespace Utilities.Feeder
{
    /// <summary>
    /// 
    /// </summary>
    public interface IXmlFeed : ICachedPropertiesProvider
    {
        /// <summary>
        /// Gets or sets the raw contents.
        /// </summary>
        /// <value>The raw contents.</value>
        string RawContents { get; set; }
    }
}
