/*
 * ParagraphExtensions.cs - Extensions for Paragraph
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

using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Wordprocessing
{
    /// <summary>
    /// 
    /// </summary>
    public static class ParagraphExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public static string GetParagraphStyleId(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphStyleId != null)
                return paragraph.ParagraphProperties.ParagraphStyleId.Val;
            else
                return "Normal";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>The character style id or null</returns>
        public static string GetCharacterStyleId(this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");

            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.ParagraphMarkRunProperties != null)
            {
                ParagraphMarkRunProperties rPr = paragraph.ParagraphProperties.ParagraphMarkRunProperties;
                RunStyle rStyle = rPr.GetFirstChild<RunStyle>();
                if (rStyle != null)
                    return rStyle.Val;
            }

            // At this point, we have not found an rStyle element.
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Style GetParagraphStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            return document.GetParagraphStyle(paragraph.GetParagraphStyleId());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Style GetCharacterStyle(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            string styleId = paragraph.GetCharacterStyleId();
            if (styleId != null)
                return document.GetCharacterStyle(styleId);
            else
                return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetOutlineLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            Style style = document.GetParagraphStyle(paragraph.GetParagraphStyleId());
            if (style.StyleParagraphProperties != null && style.StyleParagraphProperties.OutlineLevel != null)
                return style.StyleParagraphProperties.OutlineLevel.Val + 1;
            else
                return 10;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static int GetListLevel(this Paragraph paragraph, WordprocessingDocument document)
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            ParagraphProperties pPr = paragraph.ParagraphProperties;
            if (pPr != null)
            {
                // Check whether paragraph has numbering properties.
                NumberingProperties numPr = pPr.NumberingProperties;
                if (numPr != null)
                {
                    NumberingLevelReference ilvl = numPr.NumberingLevelReference;
                    if (ilvl != null)
                        return ilvl.Val + 1;
                    else
                        return 1;
                }
                else
                {
                    // Check whether the style has numbering properties.
                    Style style = paragraph.GetParagraphStyle(document);
                    StyleParagraphProperties spPr = style.StyleParagraphProperties;
                    if (spPr != null)
                    {
                        numPr = spPr.NumberingProperties;
                        if (numPr != null)
                        {
                            NumberingLevelReference ilvl = numPr.NumberingLevelReference;
                            if (ilvl != null)
                                return ilvl.Val + 1;
                            else
                                return 1;
                        }
                    }                       
                }
            }
             
            // No numbering properties found.
            return 0;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="paragraph"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static T GetParagraphPropertiesLeafElement<T>(this Paragraph paragraph, WordprocessingDocument document) where T : OpenXmlLeafElement
        {
            if (paragraph == null)
                throw new ArgumentNullException("paragraph");
            if (document == null)
                throw new ArgumentNullException("document");

            // See whether the paragraph style has it.
            Style pStyle = paragraph.GetParagraphStyle(document);
            if (pStyle.StyleRunProperties != null)
            {
                T leaf = pStyle.StyleRunProperties.GetFirstChild<T>();
                if (leaf != null)
                    return leaf;
            }

            // See whether a potential character style has it.
            Style rStyle = paragraph.GetCharacterStyle(document);
            if (rStyle != null && rStyle.StyleRunProperties != null)
            {
                T leaf = rStyle.StyleRunProperties.GetFirstChild<T>();
                if (leaf != null)
                    return leaf;
            }
                
            // See whether the paragraph's run properties have this leaf element.
            ParagraphProperties pPr = paragraph.ParagraphProperties;
            if (pPr != null && pPr.ParagraphMarkRunProperties != null)
            {
                return pPr.ParagraphMarkRunProperties.GetFirstChild<T>();
            }

            // There is no such leaf.
            return null;
        }
    }
}
