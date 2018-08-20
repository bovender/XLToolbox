/* Command.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
 */
namespace XLToolbox
{
    /// <summary>
    /// Enumeration of user-entry commands of the XL Toolbox addin.
    /// </summary>
    public enum Command
    {
        // Start with value 1 to prevent subtle bugs with YAML serialization/deserialization
        About = 1,
        CheckForUpdates,
        ThrowError,
        SheetManager,
        ExportSelection,
        ExportSelectionLast,
        BatchExport,
        BatchExportLast,
        ExportScreenshot,
        Donate,
        QuitExcel,
        OpenCsv,
        OpenCsvWithParams,
        SaveCsv,
        SaveCsvWithParams,
        SaveCsvRange,
        SaveCsvRangeWithParams,
        SaveAs,
        JumpToTarget,
        CopyPageSetup,
        SelectAllShapes,
        Anova1Way,
        AnovaRepeat,
        Anova2Way,
        FormulaBuilder,
        SelectionAssistant,
        LinearRegression,
        Correlation,
        TransposeWizard,
        MultiHisto,
        Allocate,
        LastErrorBars,
        AutomaticErrorBars,
        InteractiveErrorBars,
        ChartDesign,
        MoveDataSeriesLeft,
        MoveDataSeriesRight,
        Annotate,
        SpreadScatter,
        SeriesToFront,
        SeriesForward,
        SeriesBackward,
        SeriesToBack,
        AddSeries,
        CopyChart,
        PointChart,
        Watermark,
        UserSettings,
        LegacyPrefs,
        Shortcuts,
        Backups,
        Properties,
        OpenFromCell,
    }

    public static class CmmandExtensions
    {
        public static bool CanHaveKeyboardShortcut(this Command command)
        {
            switch (command)
            {
                case Command.OpenFromCell:
                    return false;
                default:
                    return true;
            }
        }
    }
}
