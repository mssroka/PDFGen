using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using iText.Forms;
using iText.Kernel.Pdf;

class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

public class MainForm : Form
{
    private TextBox excelFileTextBox;
    private TextBox outputDirectoryTextBox;
    private Button generateButton;

    private static readonly string PdfTemplateFileName = "Template.pdf";

    public MainForm()
    {
        this.Text = "PDF Generator";
        this.Size = new System.Drawing.Size(460, 200);

        var excelFileLabel = new Label { Text = "Ścieżka do pliku Excel:", AutoSize = true, Left = 10, Top = 20 };
        excelFileTextBox = new TextBox { Left = 150, Top = 20, Width = 200 };
        var excelFileButton = new Button { Text = "Wybierz...", Left = 360, Top = 18, Width = 80 };
        excelFileButton.Click += (sender, e) => { excelFileTextBox.Text = ChooseFile("Wybierz plik Excel", "Excel Files|*.xls;*.xlsx"); };

        var outputDirectoryLabel = new Label { Text = "Ścieżka do zapisu PDF:", AutoSize = true, Left = 10, Top = 60 };
        outputDirectoryTextBox = new TextBox { Left = 150, Top = 60, Width = 200 };
        var outputDirectoryButton = new Button { Text = "Wybierz...", Left = 360, Top = 58, Width = 80 };
        outputDirectoryButton.Click += (sender, e) => { outputDirectoryTextBox.Text = ChooseFolder("Wybierz folder do zapisu PDF"); };

        generateButton = new Button { Text = "Generuj PDF-y", Left = 150, Top = 100, Width = 200 };
        generateButton.Click += GenerateButton_Click;

        this.Controls.Add(excelFileLabel);
        this.Controls.Add(excelFileTextBox);
        this.Controls.Add(excelFileButton);
        this.Controls.Add(outputDirectoryLabel);
        this.Controls.Add(outputDirectoryTextBox);
        this.Controls.Add(outputDirectoryButton);
        this.Controls.Add(generateButton);
    }

    private void GenerateButton_Click(object sender, EventArgs e)
    {
        string pdfTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, PdfTemplateFileName);
        string excelFilePath = excelFileTextBox.Text;
        string outputDirectory = outputDirectoryTextBox.Text;

        if (!File.Exists(pdfTemplatePath))
        {
            MessageBox.Show("Szablon PDF nie został znaleziony w katalogu wyjściowym.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!File.Exists(excelFilePath))
        {
            MessageBox.Show("Podany plik Excel nie istnieje.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (!Directory.Exists(outputDirectory))
        {
            MessageBox.Show("Podany katalog nie istnieje.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var excelData = ReadExcel(excelFilePath);

        if (excelData.Count > 0)
        {
            foreach (var data in excelData)
            {
                string outputPdfPath = Path.Combine(outputDirectory, $"{data["Numer zewnetrzny"]} {data["Nazwisko odbiorcy"]}.pdf");
                try
                {
                    FillPdf(pdfTemplatePath, outputPdfPath, data);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Błąd podczas tworzenia PDF dla numeru {data["Numer zewnetrzny"]}: {ex.Message}", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            MessageBox.Show("Sukces!");
        }
        else
        {
            MessageBox.Show("Brak danych w pliku Excel.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private string ChooseFile(string title, string filter)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Title = title;
            openFileDialog.Filter = filter;
            return openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : string.Empty;
        }
    }

    private string ChooseFolder(string description)
    {
        using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
        {
            folderBrowserDialog.Description = description;
            return folderBrowserDialog.ShowDialog() == DialogResult.OK ? folderBrowserDialog.SelectedPath : string.Empty;
        }
    }

    public static List<Dictionary<string, string>> ReadExcel(string excelFilePath)
    {
        var dataList = new List<Dictionary<string, string>>();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                reader.Read(); // Skip the header row

                while (reader.Read())
                {
                    var data = new Dictionary<string, string>
                    {
                        { "Numer zewnetrzny", reader.GetValue(0)?.ToString() },
                        { "Referencja", reader.GetValue(1)?.ToString() },
                        { "Uwagi", reader.GetValue(2)?.ToString() },
                        { "Nazwisko nadawcy", reader.GetValue(3)?.ToString() },
                        { "Email Nadawcy", reader.GetValue(4)?.ToString() },
                        { "Telefon Nadawcy", reader.GetValue(5)?.ToString() },
                        { "Wartość Celna", reader.GetValue(6)?.ToString() },
                        { "Waluta wartości celnej", reader.GetValue(7)?.ToString() },
                        { "Nazwisko odbiorcy", reader.GetValue(8)?.ToString() }
                    };
                    dataList.Add(data);
                }
            }
        }
        return dataList;
    }

    public static void FillPdf(string templatePath, string outputPath, Dictionary<string, string> data)
    {
        try
        {
            using (var pdfReader = new PdfReader(templatePath))
            using (var pdfWriter = new PdfWriter(outputPath))
            using (var pdfDoc = new PdfDocument(pdfReader, pdfWriter))
            {
                var form = PdfAcroForm.GetAcroForm(pdfDoc, true);
                var fields = form.GetAllFormFields();

                if (fields.ContainsKey("tracking number")) fields["tracking number"].SetValue(data["Numer zewnetrzny"]);
                if (fields.ContainsKey("numery faktur")) fields["numery faktur"].SetValue(data["Referencja"]);
                if (fields.ContainsKey("opis towaru")) fields["opis towaru"].SetValue(data["Uwagi"]);
                if (fields.ContainsKey("ilość faktur")) fields["ilość faktur"].SetValue("1");
                if (fields.ContainsKey("imię i nazwisko wysyłającego")) fields["imię i nazwisko wysyłającego"].SetValue(data["Nazwisko nadawcy"]);
                if (fields.ContainsKey("dane kontaktowe")) fields["dane kontaktowe"].SetValue($"{data["Telefon Nadawcy"]} {data["Email Nadawcy"]}");
                if (fields.ContainsKey("waluta2")) fields["waluta2"].SetValue($"{data["Wartość Celna"]} {data["Waluta wartości celnej"]}");

                fields["Group1"].SetValue("Wybór2");
                fields["Group2"].SetValue("Wybór3");
                fields["Group3"].SetValue("Wybór7");
                fields["Group4"].SetValue("Wybór18");
                fields["Group5"].SetValue("Wybór16");

                form.FlattenFields();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Błąd podczas wypełniania PDF: {ex.Message}", ex);
        }
    }
}
