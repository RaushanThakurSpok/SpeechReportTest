using System;
using System.Collections.Generic;
using System.Text;

namespace Utilities.EcmaScript
{
    /// <summary>
    /// 
    /// </summary>
    public class EcmaScriptProgram : EcmaScriptNode
    {
        private List<EcmaScriptSourceElement> m_sourceElements;

        /// <summary>
        /// Gets or sets the source elements.
        /// </summary>
        /// <value>The source elements.</value>
        public List<EcmaScriptSourceElement> SourceElements
        {
            get
            {
                return m_sourceElements;
            }
            set
            {
                m_sourceElements = value;
            }
        }
    }
}
