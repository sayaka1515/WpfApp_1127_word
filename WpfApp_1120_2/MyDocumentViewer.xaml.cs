using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using HtmlAgilityPack;

namespace WpfApp_1120_2
{
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;
        public MyDocumentViewer()
        {
            InitializeComponent();
            backgroundColorPicker.SelectedColor = ((SolidColorBrush)documentRichTextBox.Background).Color;
            fontColorPicker.SelectedColor = fontColor;
            foreach (FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                fontFamilyComboBox.Items.Add(fontFamily.Source);
            }
            fontFamilyComboBox.SelectedIndex = 8;

            fontSizeComboBox.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            fontSizeComboBox.SelectedIndex = 2;
        }

        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML File (*.html)|*.html|All files (*.*)|*.*";
            openFileDialog.DefaultExt = ".rtf";
            openFileDialog.AddExtension = true;

            if (openFileDialog.ShowDialog() == true)
            {
                FileStream fileStream = new FileStream(openFileDialog.FileName, FileMode.Open);
                TextRange range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);
                if (openFileDialog.FilterIndex == 1) // RTF
                {
                    range.Load(fileStream, DataFormats.Rtf);
                }
                else if (openFileDialog.FilterIndex == 2) // HTML
                {
                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        string htmlText = reader.ReadToEnd();
                        string rtfText = ConvertHtmlToRtf(htmlText);
                        using (MemoryStream rtfStream = new MemoryStream(Encoding.UTF8.GetBytes(rtfText)))
                        {
                            range.Load(rtfStream, DataFormats.Rtf);
                        }
                    }
                }
                fileStream.Close();
            }
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|HTML File (*.html)|*.html|All files (*.*)|*.*";
            saveFileDialog.DefaultExt = ".rtf";
            saveFileDialog.AddExtension = true;

            if (saveFileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);
                using (FileStream fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create))
                {
                    if (saveFileDialog.FilterIndex == 1) // RTF
                    {
                        range.Save(fileStream, DataFormats.Rtf);
                    }
                    else if (saveFileDialog.FilterIndex == 2) // HTML
                    {
                        string plainText = ConvertToHtml(range);
                        Color backgroundColor = ((SolidColorBrush)documentRichTextBox.Background).Color;
                        string htmlTemplate = GenerateHtmlTemplate(plainText, backgroundColor);
                        using (StreamWriter writer = new StreamWriter(fileStream))
                        {
                            writer.Write(htmlTemplate);
                        }
                    }
                }
            }
        }

        private string ConvertToHtml(TextRange range)
        {
            MemoryStream rtfMemoryStream = new MemoryStream();
            range.Save(rtfMemoryStream, DataFormats.Rtf);
            rtfMemoryStream.Position = 0;

            TextRange plainTextRange = new TextRange(range.Start, range.End);
            string plainText = plainTextRange.Text;

            return plainText;
        }

        private string ConvertHtmlToRtf(string htmlText)
        {
            // 使用 HtmlAgilityPack 來轉換 HTML 為 RTF
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(htmlText);
            StringWriter writer = new StringWriter();
            doc.Save(writer);
            return writer.ToString();
        }


        private string GenerateHtmlTemplate(string content, Color backgroundColor)
        {
            string backgroundColorHex = $"#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}";
            return $@"
<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>Document</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: {backgroundColorHex};
        }}
    </style>
</head>
<body>
    <pre>{System.Net.WebUtility.HtmlEncode(content)}</pre>
</body>
</html>";
        }

        private void documentRichTextBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            var property_bold = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            boldToggleButton.IsChecked = (property_bold != DependencyProperty.UnsetValue) && (property_bold.Equals(FontWeights.Bold));

            var property_italic = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            italicToggleButton.IsChecked = (property_italic != DependencyProperty.UnsetValue) && (property_italic.Equals(FontStyles.Italic));

            var property_underline = documentRichTextBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            underlineToggleButton.IsChecked = (property_underline != DependencyProperty.UnsetValue) && (property_underline.Equals(TextDecorations.Underline));

            var property_fontSize = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            fontSizeComboBox.SelectedItem = property_fontSize;

            var property_fontFamily = documentRichTextBox.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            fontFamilyComboBox.SelectedItem = property_fontFamily.ToString();

            var property_fontColor = documentRichTextBox.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            fontColorPicker.SelectedColor = ((SolidColorBrush)property_fontColor).Color;
        }

        private void fontSizeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontSizeComboBox.SelectedItem != null)
            {
                documentRichTextBox.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeComboBox.SelectedItem);
            }
        }

        private void fontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontFamilyComboBox.SelectedItem != null)
            {
                documentRichTextBox.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, new FontFamily((string)fontFamilyComboBox.SelectedItem));
            }
        }

        private void fontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = fontColorPicker.SelectedColor.Value;
            SolidColorBrush brush = new SolidColorBrush(fontColor);
            documentRichTextBox.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        }

        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            documentRichTextBox.Document.Blocks.Clear();
        }

        private void backgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            if (backgroundColorPicker.SelectedColor.HasValue)
            {
                Color selectedColor = backgroundColorPicker.SelectedColor.Value;
                documentRichTextBox.Background = new SolidColorBrush(selectedColor);
            }
        }
    }
}
