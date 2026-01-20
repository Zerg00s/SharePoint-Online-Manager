namespace SharePointOnlineManager.Forms.Dialogs;

/// <summary>
/// Dialog for importing URLs from text or file.
/// </summary>
public class UrlImportDialog : Form
{
    private TextBox _urlsTextBox = null!;
    private Button _pasteButton = null!;
    private Button _loadFileButton = null!;
    private Label _countLabel = null!;

    public List<string> ImportedUrls { get; private set; } = [];

    public UrlImportDialog()
    {
        InitializeUI();
    }

    private void InitializeUI()
    {
        Text = "Import URLs";
        Size = new Size(600, 450);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var instructionLabel = new Label
        {
            Text = "Paste site URLs below (one per line) or load from a file:",
            Location = new Point(15, 15),
            AutoSize = true
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Location = new Point(15, 40),
            Size = new Size(550, 35),
            FlowDirection = FlowDirection.LeftToRight
        };

        _pasteButton = new Button
        {
            Text = "Paste from Clipboard",
            Size = new Size(140, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _pasteButton.Click += PasteButton_Click;

        _loadFileButton = new Button
        {
            Text = "Load from File...",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 20, 0)
        };
        _loadFileButton.Click += LoadFileButton_Click;

        _countLabel = new Label
        {
            Text = "0 URLs",
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };

        buttonPanel.Controls.AddRange(new Control[] { _pasteButton, _loadFileButton, _countLabel });

        _urlsTextBox = new TextBox
        {
            Location = new Point(15, 80),
            Size = new Size(550, 280),
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            Font = new Font("Consolas", 9F),
            WordWrap = false
        };
        _urlsTextBox.TextChanged += UrlsTextBox_TextChanged;

        var okButton = new Button
        {
            Text = "Import",
            DialogResult = DialogResult.OK,
            Location = new Point(405, 375),
            Size = new Size(75, 28)
        };
        okButton.Click += OkButton_Click;

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(490, 375),
            Size = new Size(75, 28)
        };

        AcceptButton = okButton;
        CancelButton = cancelButton;

        Controls.AddRange(new Control[]
        {
            instructionLabel, buttonPanel, _urlsTextBox, okButton, cancelButton
        });
    }

    private void PasteButton_Click(object? sender, EventArgs e)
    {
        if (Clipboard.ContainsText())
        {
            _urlsTextBox.Text = Clipboard.GetText();
        }
    }

    private void LoadFileButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            Title = "Select URL File"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var content = File.ReadAllText(dialog.FileName);
                _urlsTextBox.Text = content;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void UrlsTextBox_TextChanged(object? sender, EventArgs e)
    {
        UpdateCount();
    }

    private void UpdateCount()
    {
        var urls = ParseUrls(_urlsTextBox.Text);
        _countLabel.Text = $"{urls.Count} valid URLs";
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        ImportedUrls = ParseUrls(_urlsTextBox.Text);

        if (ImportedUrls.Count == 0)
        {
            MessageBox.Show("No valid URLs found.", "Validation",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
        }
    }

    private static List<string> ParseUrls(string text)
    {
        var urls = new List<string>();

        if (string.IsNullOrWhiteSpace(text))
            return urls;

        var lines = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            // Handle CSV format (URL might be in first column)
            if (trimmed.Contains(','))
            {
                trimmed = trimmed.Split(',')[0].Trim().Trim('"');
            }

            // Validate URL
            if (Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) &&
                (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps) &&
                uri.Host.Contains("sharepoint.com", StringComparison.OrdinalIgnoreCase))
            {
                // Normalize URL (remove trailing slash)
                var normalizedUrl = uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
                if (!urls.Contains(normalizedUrl, StringComparer.OrdinalIgnoreCase))
                {
                    urls.Add(normalizedUrl);
                }
            }
        }

        return urls;
    }
}
