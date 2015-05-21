using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractArchitect.OpenXml.Extensions
{
    public static class NumberingPropertiesExtensions
    {
        /// <summary>
        /// Gets the <see cref="AbstractNum" /> element that is "directly" associated with the
        /// <see cref="NumberingProperties" /> element. The AbstractNum element that is returned
        /// might just point back to a list style.
        /// </summary>
        /// <seealso cref="GetEffectiveAbstractNum" />
        /// <param name="numPr">The <see cref="NumberingProperties" /> element.</param>
        /// <param name="document">The <see cref="WordprocessingDocument" /> instance.</param>
        /// <returns>The <see cref="AbstractNum" /> element.</returns>
        public static AbstractNum GetAbstractNum(this NumberingProperties numPr, WordprocessingDocument document)
        {
            if (numPr == null) return null;

            var numIdVal = numPr.GetNumberingIdValue();
            if (numIdVal == 0) return null;

            if (document == null)
                throw new ArgumentNullException("document");

            var numbering = document.ProduceNumberingElement();

            var num = numbering.Elements<NumberingInstance>().FirstOrDefault(e => e.NumberID.Value == numIdVal);
            if (num == null) return null;

            return numbering.Elements<AbstractNum>()
                .FirstOrDefault(e => e.AbstractNumberId.Value == num.AbstractNumId.Val.Value);
        }

        /// <summary>
        /// Get the <see cref="NumberingProperties" /> element's effective <see cref="AbstractNum" /> element,
        /// following the additional level of indirection that might be introduced by a list style.
        /// </summary>
        /// <param name="numPr">The <see cref="NumberingProperties" /> element.</param>
        /// <param name="document">The <see cref="WordprocessingDocument" /> instance.</param>
        /// <returns>The <see cref="AbstractNum" /> element.</returns>
        public static AbstractNum GetEffectiveAbstractNum(this NumberingProperties numPr,
            WordprocessingDocument document)
        {
            var abstractNum = numPr.GetAbstractNum(document);
            if (abstractNum == null) return null;

            var numStyleLink = abstractNum.NumberingStyleLink;
            if (numStyleLink == null) return abstractNum;

            var numberingStyleId = numStyleLink.Val.Value;
            var styles = document.ProduceStylesElement();
            var numberingStyle = styles.Elements<Style>().FirstOrDefault(e => e.StyleId.Value == numberingStyleId);

            return numberingStyle != null ? numberingStyle.GetNumberingProperties().GetAbstractNum(document) : null;
        }

        public static int GetIndentationLeft(this NumberingProperties numPr, WordprocessingDocument document)
        {
            var abstractNum = numPr.GetEffectiveAbstractNum(document);
            if (abstractNum == null) return 0;

            var ilvlVal = numPr.GetNumberingLevelReferenceValue();
            var lvl = abstractNum.Elements<Level>().FirstOrDefault(e => e.LevelIndex.Value == ilvlVal);
            if (lvl == null) return 0;

            var pPr = lvl.PreviousParagraphProperties;
            if (pPr == null) return 0;

            var ind = pPr.Indentation;
            if (ind != null && ind.Left != null)
                return int.Parse(ind.Left.Value);

            return 0;
        }

        public static int GetNumberingIdValue(this NumberingProperties numPr)
        {
            return numPr != null && numPr.NumberingId != null
                ? numPr.NumberingId.Val.Value
                : 0;
        }

        public static int GetNumberingLevelReferenceValue(this NumberingProperties numPr)
        {
            return numPr != null && numPr.NumberingLevelReference != null
                ? numPr.NumberingLevelReference.Val.Value
                : 0;
        }

        public static string GetNumberingText(this NumberingProperties numPr, NumberingState numberingState,
            WordprocessingDocument document)
        {
            if (numberingState == null)
                throw new ArgumentNullException("numberingState");

            return numberingState.GetNumberingText(numPr, document);
        }
    }

    internal static class NumberConversionExtensions
    {
        private static readonly char[] LowerLetters =
        {
            'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
            'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
            'u', 'v', 'w', 'x', 'y', 'z'
        };

        public static string Convert(this int number, Level level)
        {
            if (level.IsLegalNumberingStyle != null && level.IsLegalNumberingStyle.Val.Value)
                return Convert(number, NumberFormatValues.Decimal);

            var numberFormat = level.NumberingFormat != null
                ? level.NumberingFormat.Val.Value
                : NumberFormatValues.Decimal;

            return Convert(number, numberFormat);
        }

        public static string Convert(this int number, NumberFormatValues numberFormat)
        {
            switch (numberFormat)
            {
                case NumberFormatValues.Decimal:
                    return number.ToString();
                case NumberFormatValues.LowerLetter:
                    return number.ToLowerLetter();
                case NumberFormatValues.UpperLetter:
                    return number.ToUpperLetter();
                case NumberFormatValues.LowerRoman:
                    return number.ToLowerRoman();
                case NumberFormatValues.UpperRoman:
                    return number.ToUpperRoman();
                default:
                    return number.ToString();
            }
        }

        public static string ToLowerLetter(this int value)
        {
            if (value < 1)
                throw new ArgumentOutOfRangeException("value");

            var index = (value - 1)%26;
            var count = (value - 1)/26 + 1;

            return new string(LowerLetters[index], count);
        }

        public static string ToLowerRoman(this int value)
        {
            return ToUpperRoman(value).ToLower();
        }

        public static string ToUpperLetter(this int value)
        {
            return value.ToLowerLetter().ToUpper();
        }

        // Source: http://stackoverflow.com/questions/7040289/converting-integers-to-roman-numerals
        public static string ToUpperRoman(this int value)
        {
            if (value < 0)
                throw new ArgumentOutOfRangeException("value");

            var sb = new StringBuilder();
            var remain = value;
            while (remain > 0)
            {
                if (remain >= 1000)
                {
                    sb.Append("M");
                    remain -= 1000;
                }
                else if (remain >= 900)
                {
                    sb.Append("CM");
                    remain -= 900;
                }
                else if (remain >= 500)
                {
                    sb.Append("D");
                    remain -= 500;
                }
                else if (remain >= 400)
                {
                    sb.Append("CD");
                    remain -= 400;
                }
                else if (remain >= 100)
                {
                    sb.Append("C");
                    remain -= 100;
                }
                else if (remain >= 90)
                {
                    sb.Append("XC");
                    remain -= 90;
                }
                else if (remain >= 50)
                {
                    sb.Append("L");
                    remain -= 50;
                }
                else if (remain >= 40)
                {
                    sb.Append("XL");
                    remain -= 40;
                }
                else if (remain >= 10)
                {
                    sb.Append("X");
                    remain -= 10;
                }
                else if (remain >= 9)
                {
                    sb.Append("IX");
                    remain -= 9;
                }
                else if (remain >= 5)
                {
                    sb.Append("V");
                    remain -= 5;
                }
                else if (remain >= 4)
                {
                    sb.Append("IV");
                    remain -= 4;
                }
                else if (remain >= 1)
                {
                    sb.Append("I");
                    remain -= 1;
                }
                else throw new Exception("Unexpected error.");
            }

            return sb.ToString();
        }
    }

    /// <summary>
    /// Provides a simplistic implementation. Will have to consider abstract numbering definitions.
    /// </summary>
    public class NumberingListState
    {
        private readonly int[] _levelRestartIndex = { 0, 1, 2, 3, 4, 5, 6, 7, 8 };
        private readonly Level[] _levels = new Level[9];
        private readonly string[] _levelTexts = new string[9];
        private readonly int[] _numbers = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        private readonly int[] _startNumberingValues = { 0, 0, 0, 0, 0, 0, 0, 0, 0 };

        /// <summary>
        /// Initializes a new <see cref="NumberingListState" /> instance.
        /// </summary>
        /// <param name="abstractNum">The <see cref="AbstractNum" /> element represented by this instance.</param>
        internal NumberingListState(AbstractNum abstractNum)
        {
            if (abstractNum == null)
                throw new ArgumentNullException("abstractNum");

            foreach (var level in abstractNum.Elements<Level>())
            {
                var levelIndex = level.LevelIndex.Value;

                _levels[levelIndex] = level;
                _levelTexts[levelIndex] = level.LevelText != null ? level.LevelText.Val.Value : String.Empty;
                _startNumberingValues[levelIndex] = level.StartNumberingValue != null
                    ? level.StartNumberingValue.Val.Value
                    : 0;
                _levelRestartIndex[levelIndex] = level.LevelRestart != null
                    ? level.LevelRestart.Val.Value
                    : levelIndex;
            }

            for (var levelIndex = 0; levelIndex < 9; levelIndex++)
            {
                _numbers[levelIndex] = _startNumberingValues[levelIndex] - 1;
            }
        }

        public string GetNumberingText(int levelIndex)
        {
            if (levelIndex < 0 || levelIndex > 8)
                throw new ArgumentOutOfRangeException("levelIndex");

            // Increment numbering up to level index before producing number.
            for (var i = 0; i < levelIndex; i++)
            {
                if (_numbers[i] < _startNumberingValues[i])
                    _numbers[i]++;
            }
            _numbers[levelIndex]++;

            // Restart lower levels
            if (levelIndex < 8) RestartLevelsBelow(levelIndex);

            var level = _levels[levelIndex];
            var isLegalNumbering = level.IsLegalNumberingStyle != null &&
                                   (level.IsLegalNumberingStyle.Val == null || level.IsLegalNumberingStyle.Val.Value);

            // Produce numbering string without suffix.
            var numberingString = _levelTexts[levelIndex];
            for (var i = 0; i <= levelIndex; i++)
            {
                var levelNumber = isLegalNumbering
                    ? _numbers[i].Convert(NumberFormatValues.Decimal)
                    : _numbers[i].Convert(_levels[i]);

                numberingString = numberingString.Replace("%" + (i + 1), levelNumber);
            }

            // Return numbering string with suffix.
            return numberingString + GetLevelSuffix(_levels[levelIndex]);
        }

        public void Restart(int levelIndex, int number)
        {
            if (levelIndex < 0 || levelIndex > 8)
                throw new ArgumentOutOfRangeException("levelIndex");

            _numbers[levelIndex] = number - 1;
            RestartLevelsBelow(levelIndex);
        }

        public void RestartLevelsBelow(int levelIndex)
        {
            if (levelIndex < 0 || levelIndex > 8)
                throw new ArgumentOutOfRangeException("levelIndex");

            for (var i = levelIndex + 1; i < 9; i++)
            {
                if (levelIndex + 1 <= _levelRestartIndex[i])
                    _numbers[i] = _startNumberingValues[i] - 1;
            }
        }

        private static string GetLevelSuffix(Level level)
        {
            if (level.LevelSuffix == null ||
                level.LevelSuffix.Val == null ||
                level.LevelSuffix.Val == LevelSuffixValues.Tab)
            {
                return "\t";
            }
            return level.LevelSuffix.Val == LevelSuffixValues.Space ? " " : String.Empty;
        }
    }

    public class NumberingState
    {
        private readonly Dictionary<AbstractNum, NumberingListState> _dictionary =
            new Dictionary<AbstractNum, NumberingListState>();

        public string GetNumberingText(NumberingProperties numPr, WordprocessingDocument document)
        {
            if (numPr == null)
                throw new ArgumentNullException("numPr");

            var numIdVal = numPr.GetNumberingIdValue();
            if (numIdVal == 0) return string.Empty;

            var numbering = document.ProduceNumberingElement();
            var num = numbering.Elements<NumberingInstance>().FirstOrDefault(e => e.NumberID.Value == numIdVal);
            if (num == null) return string.Empty;

            var abstractNum = numPr.GetEffectiveAbstractNum(document);
            if (abstractNum == null) return string.Empty;

            var numberingListState = GetOrCreateNumberingListState(abstractNum);
            var ilvlVal = numPr.GetNumberingLevelReferenceValue();

            var lvlOverride = num.Elements<LevelOverride>().FirstOrDefault(e => e.LevelIndex.Value == ilvlVal);
            if (lvlOverride != null && lvlOverride.StartOverrideNumberingValue != null)
            {
                var startOverrideVal = lvlOverride.StartOverrideNumberingValue.Val.Value;
                numberingListState.Restart(ilvlVal, startOverrideVal);
            }
            return numberingListState.GetNumberingText(ilvlVal);
        }

        internal string GetNumberingText(int levelIndex, AbstractNum abstractNum)
        {
            if (levelIndex < 0 || levelIndex > 8)
                throw new ArgumentOutOfRangeException("levelIndex");
            if (abstractNum == null)
                throw new ArgumentNullException("abstractNum");

            return GetOrCreateNumberingListState(abstractNum).GetNumberingText(levelIndex);
        }

        internal void Restart(int levelIndex, int startOverrideVal, AbstractNum abstractNum)
        {
            if (levelIndex < 0 || levelIndex > 8)
                throw new ArgumentOutOfRangeException("levelIndex");
            if (abstractNum == null)
                throw new ArgumentNullException("abstractNum");

            GetOrCreateNumberingListState(abstractNum).Restart(levelIndex, startOverrideVal);
        }

        private NumberingListState GetOrCreateNumberingListState(AbstractNum abstractNum)
        {
            NumberingListState numberingListState;
            if (_dictionary.ContainsKey(abstractNum))
            {
                numberingListState = _dictionary[abstractNum];
            }
            else
            {
                numberingListState = new NumberingListState(abstractNum);
                _dictionary.Add(abstractNum, numberingListState);
            }
            return numberingListState;
        }
    }
}