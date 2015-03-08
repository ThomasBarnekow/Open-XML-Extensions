using System;

namespace DocumentFormat.OpenXml.Transformation
{
    /// <summary>
    /// The class represents errors that occur during transforms.
    /// </summary>
    public class OpenXmlTransformationException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformationException" /> class.
        /// </summary>
        public OpenXmlTransformationException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformationException" /> class
        /// with a specified error message.
        /// </summary>
        /// <param name="message">The error message.</param>
        public OpenXmlTransformationException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTransformationException" /> class
        /// with a specified error message and a reference to the inner exception that is
        /// the cause of this exception.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The error message.</param>
        public OpenXmlTransformationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}