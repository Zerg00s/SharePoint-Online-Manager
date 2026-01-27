namespace SharePointOnlineManager.Forms;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        this.SuspendLayout();
        //
        // MainForm
        //
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(1024, 600);
        this.MinimumSize = new System.Drawing.Size(900, 640);
        this.Name = "MainForm";
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "SharePoint Online Manager";

        // Set the application icon from the embedded resource
        var exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        if (!string.IsNullOrEmpty(exePath) && System.IO.File.Exists(exePath))
        {
            this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
        }

        this.ResumeLayout(false);
    }
}
