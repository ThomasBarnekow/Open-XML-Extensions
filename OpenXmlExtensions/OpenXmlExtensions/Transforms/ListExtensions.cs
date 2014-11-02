/*
 * ListExtensions.cs - List<T> Extensions for Open XML Transforms
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
using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Transforms
{
    /// <summary>
    /// This class defines extension methods for the generic <see cref="List{T}"/> class.
    /// The specific extension methods provided by this class are helpful for transforming
    /// OpenXml packages.
    /// </summary>
    public static class ListExtensions
    {
        /// <summary>
        /// Appends an item or a collection of items to the list.
        /// </summary>
        /// <typeparam name="T">The list item type.</typeparam>
        /// <param name="list">The list.</param>
        /// <param name="itemOrCollection">A single item or a collection of items.</param>
        /// <returns>The list to which the item or list of items was added.</returns>
        public static List<T> Append<T>(this List<T> list, object itemOrCollection)
        {
            if (itemOrCollection != null)
            {
                if (itemOrCollection is T)
                    list.Add((T)itemOrCollection);
                else if (itemOrCollection is IEnumerable<T>)
                    list.AddRange((IEnumerable<T>)itemOrCollection);
                else
                    throw new ArgumentException("Illegal item type: " + itemOrCollection.GetType(), "item");
            }
            return list;
        }
    }

    public static class SetExtensions
    {
        public static HashSet<T> Append<T>(this HashSet<T> set, object itemOrCollection)
        {
            if (itemOrCollection != null)
            {
                if (itemOrCollection is T)
                    set.Add((T)itemOrCollection);
                else if (itemOrCollection is IEnumerable<T>)
                    foreach (T item in (IEnumerable<T>)itemOrCollection)
                        set.Add(item);
                else
                    throw new ArgumentException("Illegal item type: " + itemOrCollection.GetType(), "item");
            }
            return set;
        }
    }
}
