using System;
using System.Collections.Generic;
using System.Text;

namespace Utilities
{
    /// <summary>
    /// 
    /// </summary>
    public interface IMonotonicNumberGenerator : INumberGenerator
    {
    }

    /// <summary>
    /// 
    /// </summary>
    public class MonotonicNumberGenerator : NumberGenerator, IMonotonicNumberGenerator
    {
        /// <summary>
        /// Monotonic number generator in the range [0, int.Max]
        /// </summary>
        public static readonly IMonotonicNumberGenerator Default = new MonotonicNumberGenerator();

        private int m_current;
        private object m_lock = new object();

        /// <summary>
        /// Initializes a new instance of the <see cref="MonotonicNumberGenerator"/> class.
        /// </summary>
        public MonotonicNumberGenerator()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MonotonicNumberGenerator"/> class.
        /// </summary>
        /// <param name="maximum">The maximum.</param>
        public MonotonicNumberGenerator(int maximum)
            : base(maximum)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MonotonicNumberGenerator"/> class.
        /// </summary>
        /// <param name="minimum">The minimum.</param>
        /// <param name="maximum">The maximum.</param>
        public MonotonicNumberGenerator(int minimum, int maximum)
            : base(minimum, maximum)
        {
        }

        /// <summary>
        /// Gets the next number in the range [Minimum, Maximum]
        /// </summary>
        /// <returns></returns>
        public override int GetNextNumber()
        {
            int result;

            lock (m_lock)
            {
                if (m_current == m_maximum)
                {
                    result = m_current = m_minimum;
                }
                else
                {
                    result = ++m_current;
                }
            }

            return result;
        }
    }
}
