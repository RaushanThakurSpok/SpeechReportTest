using System;
using System.Collections.Generic;
using System.Text;

namespace Utilities
{
    /// <summary>
    /// Code marked with this attribute should be moved to the Utilities package.
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public class PendingUtilitiesAttribute : Attribute
    {
    }
}
