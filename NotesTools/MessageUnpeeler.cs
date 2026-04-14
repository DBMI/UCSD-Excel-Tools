using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace NotesTools
{
    public enum MessageDirection
    {
        FromPatient,
        ToPatient,
        None
    }

    internal class MessageUnpeeler
    {
        private Microsoft.Office.Interop.Excel.Application application;
        private readonly string[] DELIMITERS = { "----- Message " };
        private Range selectedColumnRng;

        private const string fromPatientPatternI = @"[tT]o:?\s?[\w,\-\s]{5,30},?\s?(?:DO|LVN|MD|ND|PA|RN)";
        private const string fromPatientPatternII = @"[tT]o:\s?(?:UCSD|Ucsd)";
        private const string fromPatientPatternIII = @"[tT]o:\s?[\w\W]{5,30}Refills";
        private const string fromPatientPatternIV = @"[tT]o:\s?Patient Medical Advice Request Message List";
        // Remove trailing words in the subject followed by long stretches of spaces.
        private const string messagePattern = @"[\w\W]+\s{8,}(?<message>[\w\W]+)";
        private const string payloadPatternI = @"[sS]ubject:\s?\w+\s*(?<message>[\w\W]+)";
        private const string payloadPatternII = @"-----\s*(?<message>[\w\W]+)";
        private const string toPatientPatternI = @"[fF]rom:?\s?[\w,\s]{5,30},\s?(?:DO|LVN|MD|ND|PA|RN)\s+Sent:";
        private const string toPatientPatternII = @"\w+\s(?:\w*\.?\s?)?[\w-]+,\s?(?:DO|LVN|MD|ND|PA|RN)";
        private Regex fromPatientRegexI;
        private Regex fromPatientRegexII;
        private Regex fromPatientRegexIII;
        private Regex fromPatientRegexIV;
        private Regex messageRegex;
        private Regex payloadRegexI;
        private Regex payloadRegexII;
        private Regex toPatientRegexI;
        private Regex toPatientRegexII;

        internal MessageUnpeeler()
        {
            application = Globals.ThisAddIn.Application;

            // Instantiate reusable Regexes.
            fromPatientRegexI = new Regex(fromPatientPatternI);
            fromPatientRegexII = new Regex(fromPatientPatternII);
            fromPatientRegexIII = new Regex(fromPatientPatternIII);
            fromPatientRegexIV = new Regex(fromPatientPatternIV);
            messageRegex = new Regex(messagePattern);
            payloadRegexI = new Regex(payloadPatternI);
            payloadRegexII = new Regex(payloadPatternII);
            toPatientRegexI = new Regex(toPatientPatternI);
            toPatientRegexII = new Regex(toPatientPatternII);
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

        private MessageDirection ParseDirectionFromColumnName(string columnName)
        {
            MessageDirection messageDirection = MessageDirection.None;

            if (!string.IsNullOrEmpty(columnName)) 
            { 
                if (columnName.ToLower().Contains("from patient"))
                {
                    messageDirection = MessageDirection.FromPatient;
                }
                else if (columnName.ToLower().Contains("to patient"))
                {
                    messageDirection = MessageDirection.ToPatient;
                }
                else
                {
                    // If we can't figure it out from the column name, ask user directly;
                    using (MessageDirectionForm form = new MessageDirectionForm())
                    {
                        var result = form.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            messageDirection = form.direction;
                        }
                    }
                }
            }

            return messageDirection;
        }

        private void ParseMessage(string sourceData, Range target, MessageDirection messageDirection)
        {
            string[] lines = sourceData.Split(DELIMITERS, StringSplitOptions.None);
            string payload;
            bool haveSeenMsgAddressedToProvider = false;

            foreach (string line in lines)
            {
                // Skip empty lines.
                if (line.Trim().Length > 0)
                {
                    Match payloadMatchI = payloadRegexI.Match(line);
                    Match payloadMatchII = payloadRegexII.Match(line);

                    if (payloadMatchI.Success && payloadMatchI.Groups.Count > 1)
                    {
                        payload = payloadMatchI.Groups["message"].Value.Trim();
                    }
                    else if (payloadMatchII.Success && payloadMatchII.Groups.Count > 1)
                    {
                        payload = payloadMatchII.Groups["message"].Value.Trim();
                    }
                    else
                    {
                        // Maybe it's a "raw" message?
                        payload = line;
                    }

                    if (payload.Length > 0)
                    {
                        Match messageMatch = messageRegex.Match(payload);

                        if (messageMatch.Success &&  messageMatch.Groups.Count > 1)
                        {
                            payload = messageMatch.Groups["message"].Value.Trim();
                        }

                        Match toPatientMatchI = toPatientRegexI.Match(line);
                        Match toPatientMatchII = toPatientRegexII.Match(line);
                        bool isToPatientMessage = toPatientMatchI.Success || toPatientMatchII.Success;

                        Match fromPatientMatchI = fromPatientRegexI.Match(line);
                        Match fromPatientMatchII = fromPatientRegexII.Match(line);
                        Match fromPatientMatchIII = fromPatientRegexIII.Match(line);
                        Match fromPatientMatchIV = fromPatientRegexIV.Match(line);
                        bool isFromPatientMessage = fromPatientMatchI.Success 
                                                 || fromPatientMatchII.Success 
                                                 || fromPatientMatchIII.Success
                                                 || fromPatientMatchIV.Success;

                        switch (messageDirection)
                        {
                            case MessageDirection.ToPatient:
                                if (!isFromPatientMessage)
                                {
                                    target.Value2 = payload;
                                    return;
                                }
                                
                                break;

                            case MessageDirection.FromPatient:

                                if (isFromPatientMessage)
                                {
                                    // Remember that we just found a REAL to-Provider msg.
                                    haveSeenMsgAddressedToProvider = true;

                                    // We want the -LAST- such line, so keep parsing.
                                    target.Value2 = payload;
                                }
                                else if (!haveSeenMsgAddressedToProvider && !isToPatientMessage)
                                {
                                    // Sometimes the "To:" field doesn't specify the MD/RN etc.,
                                    // so we're not sure if it's to the provider or not.
                                    // If we have already seen a msg explicitly sent to provider,
                                    // use that one. Otherwise, keep the last msg in the list that's
                                    // at least NOT explicitly FROM a provider.
                                    target.Value2 = payload;
                                }

                                break;

                            default:
                                break;
                        }
                    }
                }
            }
        }

        internal void Scan(Worksheet worksheet)
        {
            if (FindSelectedColumn(worksheet))
            {
                string selectedColumnName = selectedColumnRng.Value.ToString();
                MessageDirection messageDirection = ParseDirectionFromColumnName(selectedColumnName);
                
                // If we can't decipher the message direction, quit.
                if (messageDirection == MessageDirection.None)
                {
                    return;
                }

                string newColumnName = selectedColumnName + " (Extracted)";

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

                    try
                    {
                        target = (Range)worksheet.Cells[rowNumber, ditheredColumn.Column];
                        sourceData = worksheet.Cells[rowNumber, selectedColumnRng.Column].Value.ToString();
                        ParseMessage(sourceData, target, messageDirection);
                    }
                    catch (RuntimeBinderException)
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
