/*
 * OpenXmlTransformationException.cs - OpenXmlTransformation Exception
 * 
 * Copyright 2014 Thomas Barnekow
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 * Developer: Thomas Barnekow
 * Email: thomas<at/>barnekow<dot/>info
 * 
 * Version: 1.0.01
 */

using System;

namespace ContractArchitect.OpenXml.Transformation
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