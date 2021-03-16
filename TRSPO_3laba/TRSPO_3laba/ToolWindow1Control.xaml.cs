 namespace TRSPO_3laba
{
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;
    using System.Text.RegularExpressions;
    using Microsoft.VisualStudio.Shell;
    using System.Text;
    using System.IO;
    using System;
    using Microsoft.VisualStudio.TextManager.Interop;
    using System.Collections.Generic;
    using EnvDTE80;
    using EnvDTE;


    /// <summary>
    /// Interaction logic for ToolWindow1Control.
    /// </summary>
    /// 
    public class StatisticSet
    {
        public string FunctionName { get; set; }
        public string KeyWordCount { get; set; }
        public string LinesCount { get; set; }
        public string WithoutComments { get; set; }
    }

    public partial class ToolWindow1Control : UserControl
    {
        private string patternMultiComm = @"(/\*(.|(\r\n))*?\*/)";//@"(\/\*[^(\*\)]*\*\/)|(\/\*.*$)";//@"(\/\*.*\*\/)|((\/\*.*)$)";
        private string patternSingleComm = @"(/{2}((.*\\\r\n)*.*))";//@"\/{2}.*";
        private string patternDoubleQuotes = @"(""([^\\""\r\n]*(\\""|\\\r\n)*)*(""|(\r\n)))";
        private string patternSingleQuotes = @"('([^\\'\r\n]*(\\'|\\\r\n)*)*('|(\r\n)))";
        private string patternEmptyLine = @"([\n\r]\s*[\n\r])";

        private List<StatisticSet> items;
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow1Control"/> class.
        /// </summary>
        public ToolWindow1Control()
        {
            this.InitializeComponent();
            items = new List<StatisticSet>();
            Statistic.ItemsSource = items;
            GetFunction();
        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]

        private string CompareStartSymbol(Match match)
        {
            if (match.Value.StartsWith("/*"))
                return "/**/";
            else if (match.Value.StartsWith("//"))
                return "/**/";
            else if (match.Value.StartsWith("\"") || match.Value.StartsWith("'"))
                return Regex.Replace(match.Value, @"/\*\*/", "");
            return match.Value;
        }

        private string DelComment(string textFunc)
        {

            textFunc = Regex.Replace(textFunc, patternDoubleQuotes + @"|" + patternSingleQuotes + @"|" + patternSingleComm + @"|" + patternMultiComm, CompareStartSymbol, RegexOptions.Multiline);
            return textFunc;
        }

        private string DelEmptyLine(string textFunc)
        {
            textFunc = Regex.Replace(textFunc, patternEmptyLine, Environment.NewLine);
            return textFunc;
        }

        private string DelQuotes(string textFunc)
        {
            textFunc = Regex.Replace(textFunc, patternDoubleQuotes + @"|" + patternSingleQuotes, "\"\"");
            return textFunc;
        }

        private int CountKeyword(string textFunc)
        {
            string patternKeyword = @"alignas|alignof|and|and_eq|asm|auto|bitand|bitor|bool|break|case|catch|char|char16_t|char32_t|class|compl|const|constexpr|const_cast|continue|decltype|default|delete|do|double|dynamic_cast|else|enum|explicit|export|extern|false|float|for|friend|goto|if|inline|int|long|mutable|namespace|new|noexcept|not|not_eq|nullptr|operator|or|or_eq|private|protected|public|register|reinterpret_cast|return|short|signed|sizeof|static|static_assert|static_cast|struct|switch|template|this|thread_local|throw|true|try|typedef|typeid|typename|union|unsigned|using|virtual|void|volatile|wchar_t|while|xor|xor_eq";       
            return Regex.Matches(textFunc, patternKeyword).Count;
        }

        private void GetInfoAboutFunc(string textFunc)
        {
            int countLines = textFunc.Split('\n').Length;

            string textFuncWithoutComment = DelComment(textFunc);
            textFunc = DelEmptyLine(textFuncWithoutComment);

            int newCountLines = textFunc.Split('\n').Length - Regex.Matches(textFunc, @"(.*/\*\*/.*)").Count;

            //MessageBox.Show(textFunc);
            int openCurlyBracePos = textFunc.IndexOf('{');
            string nameFunc = Regex.Replace(Regex.Replace(textFunc.Substring(0, openCurlyBracePos), @"(\s*[\r\n]\s*)|(\s+)", " "), @"/\*\*/", "");

            items.Add(new StatisticSet() { FunctionName = nameFunc, KeyWordCount = CountKeyword(DelQuotes(textFuncWithoutComment)).ToString(), LinesCount = countLines.ToString(), WithoutComments = newCountLines.ToString() });
        }

        private string getFuncDeclaration(CodeElement codeElement)
        {
            Dispatcher.VerifyAccess();

            CodeFunction function = codeElement as CodeFunction;
            TextPoint start = function.GetStartPoint(vsCMPart.vsCMPartHeader);
            TextPoint finish = function.GetEndPoint(vsCMPart.vsCMPartBodyWithDelimiter);
            string fullSource = start.CreateEditPoint().GetText(finish);
            return fullSource;
        }

        private void GetFunction()
        {
            Dispatcher.VerifyAccess();
            DTE2 dte;
            try
            {
                dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
                ProjectItem item = dte.ActiveDocument.ProjectItem;
                FileCodeModel2 model = (FileCodeModel2)item.FileCodeModel;

                foreach (CodeElement codeElement in model.CodeElements)
                {
                    if (codeElement.Kind == vsCMElement.vsCMElementFunction)
                        GetInfoAboutFunc(getFuncDeclaration(codeElement));
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            items.Clear();
            GetFunction();
            Statistic.Items.Refresh();
        }

        private void Statistic_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }
    }
}