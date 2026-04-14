using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace DeidentifyTools
{
    // What did the user select in the HideThisNameForm?
    internal class SelectionResult
    {
        internal string alias;
        internal bool replace;
        internal string wordToReplace;

        internal SelectionResult(string alias, bool replace, string wordToReplace)
        {
            this.alias = alias;
            this.replace = replace;
            this.wordToReplace = wordToReplace;
        }
    }

    internal class Deidentifier
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private bool cancel = false;

        private const string dateOnlyFormat = @"dd/MM/yyyy";
        private const string dateOnlyPattern = @"\d{1,2}\/\d{1,2}\/\d{4}[\s\.](?!\d)";
        private Regex dateOnlyRegex;
        private string dateTimeFormat = @"dd/MM/yyyy hh:mm tt";
        private const string dateTimePattern = @"\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s[AP]M";
        private Regex dateTimeRegex;
        private const string drNamePattern = @"Dr\.\s[A-Z]\w+,?\s*(?:[A-Z]\.\s*)?(?:[A-Z]\w+)?";
        private Regex drNameRegex;
        private const string emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
        private Regex emailRegex;
        private string isoFormat = @"yyyy-MM-ddTHH:mm:ss";
        private const string isoPattern = @"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}";
        private Regex isoRegex;
        private const string monthDayOnlyPattern = @"\d{1,2}\/\d{1,2}(?![\/\d])";
        private Regex monthDayOnlyRegex;
        private const string monthSpelledOutPattern = @"\w{4,8}\s*\d{4}";
        private Regex monthSpelledOutRegex;
        // Capture single names, First Last and Last, First
        private const string namePattern = @"(?:[A-Z][a-z]+[\s,]*)+";
        private Regex nameRegex;
        private const string phonePattern = @"(\+\d{1,2}\s?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}";
        private Regex phoneRegex;
        private const string providerNameTitlePattern = @"[A-Z][\w-]+,?\s*[A-Z][\w-]*\.?\s*(?:[A-Z][\w-]*\.?)?,?\s*[A-Z]{2,}(?:\s[A-Z-]{2,})?";
        private Regex providerNameTitleRegex;
        private const string wordsBeforePattern = @"(?<left>(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?(?:[\d\w,\/\?\.]+\s)?)";
        private const string wordsAfterPattern = @"(?<right>(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?(?:\s[\d\w,\/\?\.]+)?)";

        private int dayOffset;
        private int monthOffset;
        private TimeSpan deltaT;

        private Range selectedColumnRng;
        private List<Range> selectedColumnsRng;

        private HiddenNames namesAndAliases;
        private List<string> namesToSkip;
        private string[] workInProgress;
        private HideOptions hideOption = HideOptions.Unknown;

        internal Deidentifier()
        {
            application = Globals.ThisAddIn.Application;
        }

        // Do we want to encode the hidden name ("<DJ9320>") or just mark <Redacted>?
        private bool AskUserHowToHideNames()
        {
            bool success = false;

            using (ChooseAOrBForm form = new ChooseAOrBForm(headline: "Do you want to encode names or just redact?",
                optionA: "Encode (like <SSN615>)", optionB: "Redact (<Redacted>)"))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (form.choice.Contains("Encode"))
                    {
                        hideOption = HideOptions.Encode;
                        success = true;
                    }
                    else if (form.choice.Contains("Redact"))
                    {
                        hideOption = HideOptions.Redact;
                        success = true;
                    }
                }
                else if (result == DialogResult.Cancel)
                {
                    // Then we're done here.
                }
            }

            return success;
        }

        // https://learn.microsoft.com/en-us/troubleshoot/developer/visualstudio/csharp/language-compilers/compute-hash-values
        private static string ByteArrayToString(byte[] arrInput)
        {
            int i;
            StringBuilder sOutput = new StringBuilder(arrInput.Length);

            for (i = 0; i < arrInput.Length; i++)
            {
                sOutput.Append(arrInput[i].ToString("X2"));
            }

            return sOutput.ToString();
        }

        private SelectionResult EncodeProviderName(string nameString,
                                          string leftContext,
                                          string rightContext)
        {
            string nameCleaned = nameString.Trim();
            string alias = nameCleaned;
            bool replace = false;

            if (namesAndAliases.HasName(nameCleaned))
            {
                alias = namesAndAliases.GetAlias(nameCleaned);
                replace = true;
            }
            else if (!namesToSkip.Contains(nameCleaned))
            {
                List<string> similarNames = namesAndAliases.FindSimilarNames(nameString);

                using (HideThisNameForm form = new HideThisNameForm(nameCleaned,
                                                                    similarNames,
                                                                    leftContext,
                                                                    rightContext,
                                                                    hideOption,
                                                                    namesAndAliases.FindSimilarNames))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.Cancel)
                    {
                        cancel = true;
                    }
                    else if (result == DialogResult.OK)
                    {
                        // What word are we replacing?
                        if (!string.IsNullOrEmpty(form.chosenName))
                        {
                            // This is a NEW name to be hidden (not the one we sent to the form).
                            nameCleaned = form.chosenName.Trim();
                        }

                        if (string.IsNullOrEmpty(form.linkedName))
                        {
                            if (hideOption == HideOptions.Encode)
                            {
                                alias = namesAndAliases.GetAlias(nameCleaned);
                            }
                            else
                            {
                                alias = MarkRedacted(nameCleaned);
                            }                            
                        }
                        else
                        {
                            // Perhaps we want to link a new name to an existing alias.
                            // Example: Already have an entry for "Dr. Able Provider" and we
                            // want the SAME alias for "Provider, Able, MD".
                            // (This includes the case where we've edited the string presented
                            // via form.chosenName.)
                            namesAndAliases.AddName(nameCleaned, form.linkedName);
                            alias = namesAndAliases.GetAlias(nameCleaned);
                        }

                        replace = true;
                    }
                    else
                    {
                        namesToSkip.Add(nameCleaned);
                    }
                }
            }

            return new SelectionResult(alias: alias, replace: replace, wordToReplace: nameCleaned);
        }

        private bool FindSelectedColumn(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnRng = Utilities.GetSelectedCol(application);

            if (selectedColumnRng is null)
            {
                // Then ask user to select one column.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: false))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        selectedColumnRng = Utilities.TopOfNamedColumn(worksheet, form.selectedColumns[0]);
                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        private bool FindSelectedColumns(Worksheet worksheet)
        {
            bool success = false;

            // Any column selected?
            selectedColumnsRng = Utilities.GetSelectedCols(application);

            if (selectedColumnsRng is null)
            {
                // Then ask user to select columns of interest.
                List<string> columnNames = Utilities.GetColumnNames(worksheet);

                using (ChooseCategoryForm form = new ChooseCategoryForm(columnNames, MultiSelect: true))
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        foreach (string selectedColumnName in form.selectedColumns)
                        {
                            Range thisRng = Utilities.TopOfNamedColumn(worksheet, selectedColumnName);
                            selectedColumnsRng.Add(thisRng);
                        }

                        success = true;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Then we're done here.
                        return success;
                    }
                }
            }
            else
            {
                success = true;
            }

            return success;
        }

        internal void GenerateHash(Worksheet worksheet)
        {
            if (FindSelectedColumns(worksheet))
            {
                // Make room for new column.
                Range hashColumn = Utilities.InsertNewColumn(range: selectedColumnsRng.Last(), newColumnName: "Scrambled Identifier", side: InsertSide.Right);

                string sourceData;
                Range target;

                using (ChooseHashLength form = new ChooseHashLength())
                {
                    var result = form.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        int hashLength = form.hashLength;
                        int numConsecutiveFailures = 0;
                        int rowNumber = 1;

                        while (true)
                        {
                            rowNumber++;

                            try 
                            {
                                sourceData = Utilities.CombineColumns(worksheet, rowNumber, selectedColumnsRng);

                                if (string.IsNullOrEmpty(sourceData)) 
                                { 
                                    numConsecutiveFailures++;

                                    if (numConsecutiveFailures >= 3)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    target = (Range)worksheet.Cells[rowNumber, hashColumn.Column];
                                    target.Value = StringToHash(sourceData, hashLength);
                                    numConsecutiveFailures = 0;
                                }
                            }
                            catch (NullReferenceException)
                            {
                                // An occasional miss is ok, but three in a row & we've run outta data.
                                numConsecutiveFailures++;

                                if (numConsecutiveFailures >= 3)
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        internal void HideNames(Worksheet worksheet)
        {
            if (!AskUserHowToHideNames()) { return; }

            // Initialize needed variables.
            namesAndAliases = new HiddenNames();
            namesToSkip = new List<string>();

            // Instantiate reusable Regexes.
            drNameRegex = new Regex(drNamePattern);
            emailRegex = new Regex(emailPattern);
            nameRegex = new Regex(namePattern);
            phoneRegex = new Regex(phonePattern);
            providerNameTitleRegex = new Regex(providerNameTitlePattern);

            if (FindSelectedColumns(worksheet))
            {
                foreach (Range col in selectedColumnsRng)
                {
                    HideNamesOneColumn(col);
                }
            }
        }

        private void HideNamesOneColumn(Range selectedCol)
        {
            string selectedColumnName = selectedCol.Value.ToString();
            string newColumnName = selectedColumnName + " (Names Hidden)";

            // Make room for new column.
            Range aliasedColumn = Utilities.InsertNewColumn(range: selectedCol,
                                                            newColumnName: newColumnName,
                                                            side: InsertSide.Right);

            Worksheet worksheet = selectedCol.Worksheet;

            Range target;
            MessageDirection messageDirection = new MessageDirection(selectedColumnName);

            int rowNumber = 1;
            int numConsecutiveFailures = 0;

            while (true)
            {
                // Allow user to cancel whenever they want.
                if (cancel) 
                {
                    application.StatusBar = "Cancelled";
                    break; 
                }

                rowNumber++;
                application.StatusBar = "Row: " + rowNumber;
                target = (Range)worksheet.Cells[rowNumber, aliasedColumn.Column];

                try
                {
                    string sourceData = worksheet.Cells[rowNumber, selectedCol.Column].Value;

                    if (sourceData is null)
                    {
                        numConsecutiveFailures++;

                        if (numConsecutiveFailures >= 3)
                        {
                            break;
                        }
                    } 
                    else
                    {            
                        // We'll only hide phone numbers and emails in messages FROM the patient.
                        if (messageDirection.Direction() == MessageDirectionEnum.FromPatient)
                        {
                            sourceData = ProcessOneRule(sourceData, phoneRegex, MarkRedacted);
                            sourceData = ProcessOneRule(sourceData, emailRegex, MarkRedacted);
                        }

                        sourceData = ProcessOneRuleWithGUI(sourceData, providerNameTitleRegex);
                        sourceData = ProcessOneRuleWithGUI(sourceData, drNameRegex);
                        target.Value = ProcessOneRuleWithGUI(sourceData, nameRegex);
                    }
                }
                catch (NullReferenceException)
                {
                    // An occasional miss is ok, but three in a row & we've run outta data.
                    numConsecutiveFailures++;

                    if (numConsecutiveFailures >= 3)
                    {
                        break;
                    }
                }
            }

            application.StatusBar = "Complete.";
        }

        private string MarkRedacted(string dateString)
        {
            return "<Redacted>";
        }

        internal void ObscureDateTime(Worksheet worksheet)
        {
            // Instantiate random number generator and random day, time offsets.
            Random rnd = new Random();
            dayOffset = rnd.Next(-7, 7);
            monthOffset = rnd.Next(-2, 2);
            int hourOffset = rnd.Next(-3, 3);
            int minuteOffset = rnd.Next(-20, 20);
            deltaT = new TimeSpan(hourOffset, minuteOffset, 0);

            // Instantiate reusable Regexes.
            dateOnlyRegex = new Regex(dateOnlyPattern);
            dateTimeRegex = new Regex(dateTimePattern);
            isoRegex = new Regex(isoPattern);
            monthDayOnlyRegex = new Regex(monthDayOnlyPattern);
            monthSpelledOutRegex = new Regex(monthSpelledOutPattern);

            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                string newColumnName = selectedColumnName + " (Date/Time Altered)";

                // Make room for new column.
                Range ditheredColumn = Utilities.InsertNewColumn(range: selectedColumnRng,
                                                                 newColumnName: newColumnName,
                                                                 side: InsertSide.Right);

                string sourceData;
                Range target;

                int rowNumber = 1;
                int numConsecutiveFailures = 0;

                while (true)
                {
                    rowNumber++;
                    target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];

                    try
                    {
                        sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;

                        if (sourceData is null) 
                        {
                            numConsecutiveFailures++;

                            if (numConsecutiveFailures >= 3) { break; }
                        }
                        else
                        {
                            // Modify & stuff into target cell.
                            if (!string.IsNullOrEmpty(sourceData))
                            {
                                sourceData = ProcessOneRule(sourceData, dateTimeRegex, TweakDateTime);
                                sourceData = ProcessOneRule(sourceData, dateOnlyRegex, TweakDateOnly);
                                sourceData = ProcessOneRule(sourceData, isoRegex, TweakIsoDateTime);
                                sourceData = ProcessOneRule(sourceData, monthDayOnlyRegex, TweakMonthDay);
                                sourceData = ProcessOneRule(sourceData, monthSpelledOutRegex, TweakMonthSpelledOut);
                            }

                            target.Value = sourceData;
                        }
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // It's ALREADY a DateTime object--no need to parse using RegEx.
                        DateTime sourceDateTime = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value;
                        target.Value = TweakDateTime(sourceDateTime);
                    }
                    catch (NullReferenceException)
                    {
                        // An occasional miss is ok, but three in a row & we've run outta data.
                        numConsecutiveFailures++;

                        if (numConsecutiveFailures >= 3)
                        {
                            break;
                        }
                    }
                }
            }
        }

        private string ProcessOneRule(string sourceData, Regex regex, Func<string, string> convert)
        {
            Match match = regex.Match(sourceData);
            string targetData = string.Empty;

            while (match.Success)
            {
                string beforeMatch = sourceData.Substring(0, match.Index);
                targetData += beforeMatch;
                string matchedWord = match.Value.ToString();

                targetData += convert(matchedWord);

                // Trim to just what's AFTER the match.
                sourceData = sourceData.Substring(match.Index + match.Value.Length);
                match = regex.Match(sourceData);
            }

            // Append whatever's left over.
            targetData += sourceData;

            return targetData;
        }

        private string ProcessOneRuleWithGUI(string sourceData, Regex regex)
        {
            Match match = regex.Match(sourceData);
            string targetData = string.Empty;

            while (match.Success)
            {
                string matchedWord = match.Value.ToString().Trim().TrimEnd(',');

                // Discard stuff that typically clutters up names.
                matchedWord = Regex.Replace(matchedWord, @"^Best,?", "");
                matchedWord = Regex.Replace(matchedWord, @"^Dear,?", "");
                matchedWord = Regex.Replace(matchedWord, @"^Hello,?", "");
                matchedWord = Regex.Replace(matchedWord, @"^Hi,?", "");
                matchedWord = Regex.Replace(matchedWord, @"^Sincerely,?", "");
                matchedWord = Regex.Replace(matchedWord, @"^Thanks,?", "");

                matchedWord = Regex.Replace(matchedWord, @"Best$", "");
                matchedWord = Regex.Replace(matchedWord, @"Good$", "");
                matchedWord = Regex.Replace(matchedWord, @"Thank$", "");
                matchedWord = Regex.Replace(matchedWord, @"Thanks$", "").Trim().TrimEnd(',');
                matchedWord = matchedWord.Trim().TrimStart(',').TrimEnd(',');

                if (string.IsNullOrEmpty(matchedWord))
                {
                    break;
                }

                string concatenatedPattern = wordsBeforePattern + matchedWord + wordsAfterPattern;
                Regex contextRegex = new Regex(concatenatedPattern);
                Match contextMatch = contextRegex.Match(sourceData);
                string leftContext = string.Empty;
                string rightContext = string.Empty;

                if (contextMatch.Success && contextMatch.Groups.Count > 2)
                {
                    leftContext = contextMatch.Groups["left"].Value.ToString();
                    rightContext = contextMatch.Groups["right"].Value.ToString();
                }

                SelectionResult selectionResult = EncodeProviderName(matchedWord, leftContext, rightContext);

                if (cancel) { break; }

                // Where's the detected word in the text?
                int index = sourceData.IndexOf(selectionResult.wordToReplace);
                string beforeText = sourceData.Substring(0, index);

                // Build the output.
                targetData += beforeText;

                // Apply the user's decision.
                if (selectionResult.replace)
                {
                    // Replace the word.
                    targetData += selectionResult.alias;
                }
                else
                {
                    // We're not replacing it.
                    targetData += selectionResult.wordToReplace;
                }

                // Trim to just what's AFTER the match.
                sourceData = sourceData.Substring(index + selectionResult.wordToReplace.Length);

                // Rerun the rule on rest of the data.
                match = regex.Match(sourceData);
            }

            // Append whatever's left over.
            targetData += sourceData;

            return targetData;
        }

        private string StringToHash(string sourceData, int hashLength = 32)
        {
            string hashString = string.Empty;

            // Create a byte array from source data.
            byte[] tmpSource = ASCIIEncoding.ASCII.GetBytes(sourceData);

            // Initialize a SHA256 hash object.
            using (SHA256 mySHA256 = SHA256.Create())
            {
                byte[] tmpHash = mySHA256.ComputeHash(tmpSource);
                hashString = ByteArrayToString(tmpHash);
            }

            // Return just 12 char of hash.
            return hashString.Substring(0, hashLength);
        }

        private string TweakDateOnly(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            string convertedDateString = payloadTweaked.ToString("M/d/yyyy");

            // Special case: did the Regex absorb a trailing period or space?
            if (dateString.EndsWith("."))
            {
                convertedDateString += ".";
            }

            if (dateString.EndsWith(" "))
            {
                convertedDateString += " ";
            }

            return convertedDateString;
        }

        private string TweakDateTime(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset) + deltaT;
            return payloadTweaked.ToString("M/d/yyyy h:mm tt");
        }

        private string TweakDateTime(DateTime dateTime)
        {
            DateTime payloadTweaked = dateTime.AddDays(dayOffset) + deltaT;
            return payloadTweaked.ToString("M/d/yyyy h:mm tt");
        }
        private string TweakIsoDateTime(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset) + deltaT;
            return payloadTweaked.ToString("yyyy-MM-ddTHH:mm:ss");
        }

        private string TweakMonthDay(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddDays(dayOffset);
            return payloadTweaked.ToString("M/d");
        }

        private string TweakMonthSpelledOut(string dateString)
        {
            DateTime payload = DateTime.Parse(dateString.Trim());
            DateTime payloadTweaked = payload.AddMonths(monthOffset);
            return payloadTweaked.ToString("MMMM yyyy");
        }
    }
}