using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Text.RegularExpressions;
using PowerStowToExcelConverter.Core;

namespace PowerStowToExcelConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                MessageBox.Show(ex.StackTrace.ToString());
            }

            // Loading the Translation.
            try
            {
                Controller.Instance.loadTranslator();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btn_close(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_browse(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Text documents (.txt)|*.txt";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                TextBoxBrowse.Text = filename;

                // Move the cursor to the end of the text
                TextBoxBrowse.Focus();
                TextBoxBrowse.CaretIndex = TextBoxBrowse.Text.Length;
            }
        }

        private void btn_save_as(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Documents (.xlsx)|*.xlsx";

            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                string filename = dlg.FileName;
                TextBoxSaveAs.Text = filename;

                // Focus the box and move the cursor to the end of the text
                TextBoxSaveAs.Focus();
                TextBoxSaveAs.CaretIndex = TextBoxSaveAs.Text.Length;

            }
        }

        private void TextBoxBrowse_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBoxBrowse.Text != "" && TextBoxSaveAs.Text != "")
            {
                ButtonConvert.IsEnabled = true;
            }
            else
            {
                ButtonConvert.IsEnabled = false;
            }
        }

        private void TextBoxSaveTo_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBoxBrowse.Text != "" && TextBoxSaveAs.Text != "")
            {
                ButtonConvert.IsEnabled = true;
            }
            else
            {
                ButtonConvert.IsEnabled = false;
            }
        }

        private void TextBoxAdditionalOptions_KeyDown(object sender, KeyEventArgs e)
        {
            // Only allow keys A-z and , in the input
            if ((e.Key < Key.A) || (e.Key > Key.Z) && (e.Key != Key.OemComma))
                e.Handled = true;
        }

        private void TextBoxAdditionalOptions_TextChanged(object sender, TextChangedEventArgs e)
        {
            int maxLength = 48;

            // Only allow certain amount of characters in the textbox
            if (TextBoxAdditionalOptions.Text.Length > maxLength)
            {
                TextBoxAdditionalOptions.Text = TextBoxAdditionalOptions.Text.Substring(0, maxLength);
                TextBoxAdditionalOptions.CaretIndex = TextBoxAdditionalOptions.Text.Length;
            }
        }

        private void ButtonConvert_Click(object sender, RoutedEventArgs e)
        {
            // Make the text/button unclickable
            TextBoxBrowse.IsEnabled = false;
            TextBoxSaveAs.IsEnabled = false;
            ButtonConvert.IsEnabled = false;
            TextBoxAdditionalOptions.IsEnabled = false;

            // Make sure that the input path has the correct file format. If not correct it
            if (!TextBoxBrowse.Text.Contains(@".txt"))
                TextBoxBrowse.Text = TextBoxBrowse.Text + ".txt";

            if (!TextBoxSaveAs.Text.Contains(@".xlsx"))
                TextBoxSaveAs.Text = TextBoxSaveAs.Text + ".xlsx";

            try
            {
                Controller.Instance.readFile(TextBoxBrowse.Text, TextBoxAdditionalOptions.Text);
                Controller.Instance.writeFile(TextBoxSaveAs.Text);

                // Converted successfully
                MessageBoxResult result = MessageBox.Show("Warning! Reefer containers are combined with the dry containers due to the limitations of the input file. Make sure to correct the output accordingly.\n\n" +
                                                          "Would you like to view the converted file?", "Converted Successfully!", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        // Append "path" in case the path has a space in it.
                        System.Diagnostics.Process.Start(@"excel.exe",  "\"" + TextBoxSaveAs.Text + "\"");
                    break;
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                MessageBox.Show(ex.ToString());
            }

            // Make the text/button clickable
            TextBoxBrowse.IsEnabled = true;
            TextBoxSaveAs.IsEnabled = true;
            ButtonConvert.IsEnabled = true;
            TextBoxAdditionalOptions.IsEnabled = true;
        }

    }
}
