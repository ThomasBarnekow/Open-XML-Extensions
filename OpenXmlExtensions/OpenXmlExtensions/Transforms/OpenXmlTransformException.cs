﻿/*
 * OpenXmlTransformException.cs - Exception for Open XML Transforms
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
using System.Linq;
using System.Text;

namespace DocumentFormat.OpenXml.Transforms
{
    public class OpenXmlTransformException : Exception
    {
        public OpenXmlTransformException()
            : base()
        { }

        public OpenXmlTransformException(string message)
            : base(message)
        { }

        public OpenXmlTransformException(string message, Exception innerException)
            : base(message, innerException)
        { }
    }
}