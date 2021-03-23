namespace PrismTaskPanes.Host
{
    partial class PrismTaskPanesHost
    {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = default;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuspendLayout();

            components = new System.ComponentModel.Container();

            this.elementHost = new System.Windows.Forms.Integration.ElementHost();
            this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost.Location = new System.Drawing.Point(0, 0);
            this.elementHost.Name = "elementHost";
            this.elementHost.Size = new System.Drawing.Size(150, 150);
            this.elementHost.TabIndex = 0;
            this.elementHost.Text = "elementHost";

            this.Controls.Add(this.elementHost);

            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;

            this.Name = "PrismTaskPanesHost";
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Integration.ElementHost elementHost;

        #endregion
    }
}
