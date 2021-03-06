﻿/*
 * OpenXmlElementExtensions.cs - Extensions for OpenXmlElement
 *
 * Copyright 2014-2015 Thomas Barnekow
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
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
{
    /// <summary>
    /// Provides extension methods for <see cref="OpenXmlElement" /> class.
    /// </summary>
    public static class OpenXmlElementExtensions
    {
        /// <summary>
        /// Transforms the elements (children) of the given element, applying the given transformation.
        /// </summary>
        /// <remarks>
        /// This method uses the Linq to XML Aggregate extension method in conjunction with
        /// the <see cref="ListExtensions.Append{T}" /> extension method to support transforms
        /// producing generic <see cref="IEnumerable{OpenXmlElement}" /> collections as a result.
        /// For each descendant, this incurs the overhead of creating a new
        /// <see cref="List{OpenXmlElement}" />. If it is guaranteed that transforms only return
        /// single <see cref="OpenXmlElement" /> instances or null, that overhead can be avoided
        /// by using the <see cref="SelectResultsOf" /> extension method.
        /// </remarks>
        /// <param name="elements">The collection of elements to be transformed.</param>
        /// <param name="transformation">The transformation to be applied.</param>
        /// <returns>The collection of transformed elements.</returns>
        /// <seealso cref="SelectResultsOf" />
        public static IEnumerable<OpenXmlElement> AggregateResultsOf(this IEnumerable<OpenXmlElement> elements,
            Func<OpenXmlElement, object> transformation)
        {
            return elements.Aggregate(new List<OpenXmlElement>(), (list, e) => list.Append(transformation(e)));
        }

        /// <summary>
        /// Gets or creates the element's first child of type T.
        /// </summary>
        /// <typeparam name="T">A subclass of OpenXmlElement</typeparam>
        /// <param name="element">The element</param>
        /// <returns>The existing or newly created first child of this element</returns>
        public static T Produce<T>(this OpenXmlElement element) where T : OpenXmlElement, new()
        {
            var child = element.GetFirstChild<T>();
            if (child != null) return child;

            child = new T();
            element.AppendChild(child);
            return child;
        }

        /// <summary>
        /// Transforms the elements (children) of the given element, applying the given transformation.
        /// </summary>
        /// <remarks>
        /// This method uses the Linq to XML Select extension method to create the generic
        /// collection of <see cref="OpenXmlElement" /> instances to be appended to the root
        /// element. With the current Open XML SDK, this does not work in case the result of
        /// a single transformation is a collection of <see cref="OpenXmlElement" /> instances.
        /// If it is not guaranteed that only single <see cref="OpenXmlElement" /> instances
        /// or null are returned by the transformation, <see cref="AggregateResultsOf" /> should
        /// be used instead of this method.
        /// </remarks>
        /// <param name="elements">The collection of elements to be transformed.</param>
        /// <param name="transformation">The transformation to be applied.</param>
        /// <returns>The collection of transformed elements.</returns>
        /// <seealso cref="AggregateResultsOf" />
        public static IEnumerable<OpenXmlElement> SelectResultsOf(this IEnumerable<OpenXmlElement> elements,
            Func<OpenXmlElement, object> transformation)
        {
            return elements.Select(e => (OpenXmlElement) transformation(e));
        }

        /// <summary>
        /// Sets the element's first child of a given type, either replacing an existing
        /// one or appending the first one.
        /// </summary>
        /// <typeparam name="T">A subclass of OpenXmlElement</typeparam>
        /// <param name="element">The element</param>
        /// <param name="newChild">The new child</param>
        /// <returns>The new child</returns>
        public static T SetFirstChild<T>(this OpenXmlElement element, T newChild) where T : OpenXmlElement
        {
            var oldChild = element.GetFirstChild<T>();
            return oldChild != null ? element.ReplaceChild(newChild, oldChild) : element.AppendChild(newChild);
        }

        public static XElement ToXElement(this OpenXmlElement element)
        {
            return element != null ? XElement.Parse(element.OuterXml) : null;
        }

        /// <summary>
        /// Performs a transformation of the given element and its descendants, creating a shallow
        /// clone of the element and applying the given transformation to its elements.
        /// </summary>
        /// <typeparam name="T"><see cref="OpenXmlElement" /> or a subclass thereof.</typeparam>
        /// <param name="element">The element to be transformed.</param>
        /// <param name="transformation">The transformation to be applied to the element's descendants.</param>
        /// <returns>The transformed element.</returns>
        /// <seealso cref="TransformSelecting{T}" />
        /// <seealso cref="AggregateResultsOf" />
        public static T TransformAggregating<T>(this T element, Func<OpenXmlElement, object> transformation)
            where T : OpenXmlElement
        {
            var transformedElement = (T) element.CloneNode(false);
            transformedElement.Append(element.Elements().AggregateResultsOf(transformation));
            return transformedElement;
        }

        /// <summary>
        /// Performs a transformation of the given element and its descendants, creating a shallow
        /// clone of the element and applying the given transformation to its elements.
        /// </summary>
        /// <typeparam name="T"><see cref="OpenXmlElement" /> or a subclass thereof.</typeparam>
        /// <param name="element">The element to be transformed.</param>
        /// <param name="transformation">The transformation to be applied to the element's descendants.</param>
        /// <returns>The transformed element.</returns>
        /// <seealso cref="TransformAggregating{T}" />
        /// <seealso cref="SelectResultsOf" />
        public static T TransformSelecting<T>(this T element, Func<OpenXmlElement, object> transformation)
            where T : OpenXmlElement
        {
            var transformedElement = (T) element.CloneNode(false);
            transformedElement.Append(element.Elements().SelectResultsOf(transformation));
            return transformedElement;
        }

        #region Logical Children Content
        // Source: http://blogs.msdn.com/b/ericwhite/archive/2010/01/12/logicalchildrencontent-axis-methods-open-xml-sdk-v2-strongly-typed-object-model.aspx

        public static IEnumerable<OpenXmlElement> DescendantsTrimmed(this OpenXmlElement element, Type trimType)
        {
            return DescendantsTrimmed(element, e => e.GetType() == trimType);
        }

        private static IEnumerable<OpenXmlElement> DescendantsTrimmed(this OpenXmlElement element,
            Func<OpenXmlElement, bool> predicate)
        {
            var iteratorStack = new Stack<IEnumerator<OpenXmlElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                while (iteratorStack.Peek().MoveNext())
                {
                    var currentOpenXmlElement = iteratorStack.Peek().Current;
                    if (predicate(currentOpenXmlElement))
                    {
                        yield return currentOpenXmlElement;
                        continue;
                    }
                    yield return currentOpenXmlElement;
                    iteratorStack.Push(currentOpenXmlElement.Elements().GetEnumerator());
                }
                iteratorStack.Pop();
            }
        }

        private static readonly Type[] SubRunLevelContent =
        {
            typeof (Break),
            typeof (CarriageReturn),
            typeof (DayLong),
            typeof (DayShort),
            typeof (Drawing),
            typeof (MonthLong),
            typeof (MonthShort),
            typeof (NoBreakHyphen),
            typeof (PageNumber),
            typeof (Picture),
            typeof (PositionalTab),
            typeof (SoftHyphen),
            typeof (SymbolChar),
            typeof (TabChar),
            typeof (Text),
            typeof (YearLong),
            typeof (YearShort),
            typeof (AlternateContent)
        };

        public static IEnumerable<OpenXmlElement> LogicalChildrenContent(this OpenXmlElement element)
        {
            if (element is Document)
                return element.Descendants<Body>().Take(1);

            if (element is Body ||
                element is TableCell ||
                element is TextBoxContent)
                return element
                    .DescendantsTrimmed(e =>
                        e is Table ||
                        e is Paragraph)
                    .Where(e =>
                        e is Paragraph ||
                        e is Table);

            if (element is Table)
                return element
                    .DescendantsTrimmed(typeof (TableRow))
                    .Where(e => e is TableRow);

            if (element is TableRow)
                return element
                    .DescendantsTrimmed(typeof (TableCell))
                    .Where(e => e is TableCell);

            if (element is Paragraph)
                return element
                    .DescendantsTrimmed(e => e is Run ||
                                             e is Picture ||
                                             e is Drawing)
                    .Where(e => e is Run ||
                                e is Picture ||
                                e is Drawing);

            if (element is Run)
                return element
                    .DescendantsTrimmed(e => SubRunLevelContent.Contains(e.GetType()))
                    .Where(e => SubRunLevelContent.Contains(e.GetType()));

            if (element is AlternateContent)
                return element
                    .DescendantsTrimmed(e =>
                        e is Picture ||
                        e is Drawing ||
                        e is AlternateContentFallback)
                    .Where(e =>
                        e is Picture ||
                        e is Drawing);

            if (element is Picture || element is Drawing)
                return element
                    .DescendantsTrimmed(typeof (TextBoxContent))
                    .Where(e => e is TextBoxContent);

            return new OpenXmlElement[] { };
        }

        public static IEnumerable<OpenXmlElement> LogicalChildrenContent(this IEnumerable<OpenXmlElement> source)
        {
            return source.SelectMany(e1 => e1.LogicalChildrenContent());
        }

        public static IEnumerable<OpenXmlElement> LogicalChildrenContent(this OpenXmlElement element,
            Type typeName)
        {
            return element.LogicalChildrenContent().Where(e => e.GetType() == typeName);
        }

        public static IEnumerable<OpenXmlElement> LogicalChildrenContent(this IEnumerable<OpenXmlElement> source,
            Type typeName)
        {
            return source.SelectMany(e => e.LogicalChildrenContent(typeName));
        }

        #endregion
    }
}