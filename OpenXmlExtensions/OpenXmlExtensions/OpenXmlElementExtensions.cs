/*
 * OpenXmlElementExtensions.cs - Extensions for OpenXmlElement
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

namespace DocumentFormat.OpenXml
{
    /// <summary>
    /// Provides extension methods for <see cref="OpenXmlElement"/> class.
    /// </summary>
    public static class OpenXmlElementExtensions
    {        
        /// <summary>
        /// Gets or creates the element's first child of type T.
        /// </summary>
        /// <typeparam name="T">A subclass of OpenXmlElement</typeparam>
        /// <param name="element">The element</param>
        /// <returns>The existing or newly created first child of this element</returns>
        public static T Produce<T>(this OpenXmlElement element) where T : OpenXmlElement, new()
        {
            T child = element.GetFirstChild<T>();
            if (child == null)
            {
                child = new T();
                element.AppendChild<T>(child);
            }
            return child;
        }

        /// <summary>
        /// Sets the element's first child, either replacing an existing one or appending the first one.
        /// </summary>
        /// <typeparam name="T">A subclass of OpenXmlElement</typeparam>
        /// <param name="element">The element</param>
        /// <param name="newChild">The new child</param>
        /// <returns>The new child</returns>
        public static T SetFirstChild<T>(this OpenXmlElement element, T newChild) where T : OpenXmlElement
        {
            T oldChild = element.GetFirstChild<T>();
            if (oldChild != null)
                return element.ReplaceChild<T>(newChild, oldChild);
            else
                return element.AppendChild<T>(newChild);
        }
    }
}
