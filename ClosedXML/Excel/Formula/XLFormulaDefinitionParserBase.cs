﻿// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel
{
    internal abstract class XLFormulaDefinitionParserBase
    {
        #region Public Methods

        public Tuple<string[], IXLReference[]> Parse(string formula)
        {
            if (String.IsNullOrWhiteSpace(formula))
            {
                return new Tuple<string[], IXLReference[]>(new[] {string.Empty}, new IXLReference[0]);
            }

            var formulaChunks = new List<string>();
            var references = new List<IXLReference>();
            var lastIndex = 0;


            foreach (var match in ReferenceRegex.Matches(formula).Cast<Match>())
            {
                var matchString = match.Value;
                var matchIndex = match.Index;
                if (formula.Substring(0, matchIndex).CharCount('"') % 2 == 0
                    && formula.Substring(0, matchIndex).CharCount('\'') % 2 == 0)
                {
                    // Check if the match is in between quotes
                    formulaChunks.Add(formula.Substring(lastIndex, matchIndex - lastIndex));
                    references.Add(ParseReference(matchString));
                }
                else
                    formulaChunks.Add(formula.Substring(lastIndex,
                        matchIndex - lastIndex + matchString.Length));

                lastIndex = matchIndex + matchString.Length;
            }

            if (lastIndex < formula.Length)
                formulaChunks.Add(formula.Substring(lastIndex));
            else
                formulaChunks.Add(string.Empty);

            return new Tuple<string[], IXLReference[]>(formulaChunks.ToArray(), references.ToArray());
        }

        #endregion Public Methods


        #region Protected Properties

        protected abstract Regex ReferenceRegex { get; }

        #endregion Protected Properties

        #region Protected Methods

        protected abstract IXLSimpleReference ParseSimpleReference(string simpleReferenceString);

        #endregion Protected Methods

        #region Private Methods

        private IXLCompoundReference ParseCompoundReference(string compoundReferenceString)
        {
            var part1 = compoundReferenceString.Substring(0, compoundReferenceString.IndexOf(":"));
            var part2 = compoundReferenceString.Substring(compoundReferenceString.IndexOf(":") + 1);

            var reference1 = ParseSimpleReference(part1);
            var reference2 = ParseSimpleReference(part2);

            switch (reference1)
            {
                case XLCellReference cellReference1:
                    if (!(reference2 is XLCellReference cellReference2))
                        throw new InvalidOperationException(
                            "Both parts of the address must be of type XLCellReference");
                    return new XLRangeReference(cellReference1, cellReference2);

                case XLRowReference rowReference1:
                    if (!(reference2 is XLRowReference rowReference2))
                        throw new InvalidOperationException(
                            "Both parts of the address must be of type XLRowReference");

                    return new XLRowRangeReference(rowReference1, rowReference2);

                case XLColumnReference columnReference1:
                    if (!(reference2 is XLColumnReference columnReference2))
                        throw new InvalidOperationException(
                            "Both parts of the address must be of type XLColumnReference");
                    return new XLColumnRangeReference(columnReference1, columnReference2);

                default:
                    throw new NotImplementedException($"Type {reference1.GetType()} is not supported");
            }

        }

        private IXLReference ParseReference(string referenceString)
        {
            if (referenceString.Contains(":"))
            {
                return ParseCompoundReference(referenceString);
            }

            return ParseSimpleReference(referenceString);
        }

        #endregion Private Methods
    }
}
