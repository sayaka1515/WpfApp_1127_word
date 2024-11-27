using System;
using System.Runtime.Serialization;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;

namespace WpfApp_1120_2
{
    /// <summary>
    /// MyDocumentViewer.xaml 的互動邏輯
    /// </summary>
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;
        public MyDocumentViewer()
        {
            InitializeComponent();
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
            openFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*";
            openFileDialog.DefaultExt = ".rtf";
            openFileDialog.AddExtension = true;

            if (openFileDialog.ShowDialog() == true)
            {
                FileStream fileStream = new FileStream(openFileDialog.FileName, FileMode.Open);
                TextRange range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);
                range.Load(fileStream, DataFormats.Rtf);
                fileStream.Close();
            }
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*";
            saveFileDialog.DefaultExt = ".rtf";
            saveFileDialog.AddExtension = true;

            if (saveFileDialog.ShowDialog() == true)
            {
                TextRange range;
                FileStream fileStream;
                range = new TextRange(documentRichTextBox.Document.ContentStart, documentRichTextBox.Document.ContentEnd);
                fileStream = new FileStream(saveFileDialog.FileName, FileMode.Create);
                range.Save(fileStream, DataFormats.Rtf);
                fileStream.Close();
            }
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
    }
}