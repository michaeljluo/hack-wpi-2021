
namespace Goat_Therapy_WPI_Hackathon
{
    partial class Home_Page
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Tab_Sytem = new System.Windows.Forms.TabControl();
            this.Nav_Home = new System.Windows.Forms.TabPage();
            this.Nav_Events = new System.Windows.Forms.TabPage();
            this.Nav_Resources = new System.Windows.Forms.TabPage();
            this.Nav_Survey = new System.Windows.Forms.TabPage();
            this.Tab_Sytem.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab_Sytem
            // 
            this.Tab_Sytem.AccessibleName = "";
            this.Tab_Sytem.Controls.Add(this.Nav_Home);
            this.Tab_Sytem.Controls.Add(this.Nav_Events);
            this.Tab_Sytem.Controls.Add(this.Nav_Resources);
            this.Tab_Sytem.Controls.Add(this.Nav_Survey);
            this.Tab_Sytem.Location = new System.Drawing.Point(2, 2);
            this.Tab_Sytem.Name = "Tab_Sytem";
            this.Tab_Sytem.SelectedIndex = 0;
            this.Tab_Sytem.Size = new System.Drawing.Size(257, 394);
            this.Tab_Sytem.TabIndex = 0;
            // 
            // Nav_Home
            // 
            this.Nav_Home.Location = new System.Drawing.Point(4, 22);
            this.Nav_Home.Name = "Nav_Home";
            this.Nav_Home.Padding = new System.Windows.Forms.Padding(3);
            this.Nav_Home.Size = new System.Drawing.Size(249, 368);
            this.Nav_Home.TabIndex = 0;
            this.Nav_Home.Text = "Home";
            this.Nav_Home.UseVisualStyleBackColor = true;
            // 
            // Nav_Events
            // 
            this.Nav_Events.Location = new System.Drawing.Point(4, 22);
            this.Nav_Events.Name = "Nav_Events";
            this.Nav_Events.Padding = new System.Windows.Forms.Padding(3);
            this.Nav_Events.Size = new System.Drawing.Size(249, 368);
            this.Nav_Events.TabIndex = 1;
            this.Nav_Events.Text = "Events";
            this.Nav_Events.UseVisualStyleBackColor = true;
            // 
            // Nav_Resources
            // 
            this.Nav_Resources.Location = new System.Drawing.Point(4, 22);
            this.Nav_Resources.Name = "Nav_Resources";
            this.Nav_Resources.Padding = new System.Windows.Forms.Padding(3);
            this.Nav_Resources.Size = new System.Drawing.Size(249, 368);
            this.Nav_Resources.TabIndex = 2;
            this.Nav_Resources.Text = "Resources";
            this.Nav_Resources.UseVisualStyleBackColor = true;
            // 
            // Nav_Survey
            // 
            this.Nav_Survey.Location = new System.Drawing.Point(4, 22);
            this.Nav_Survey.Name = "Nav_Survey";
            this.Nav_Survey.Padding = new System.Windows.Forms.Padding(3);
            this.Nav_Survey.Size = new System.Drawing.Size(249, 368);
            this.Nav_Survey.TabIndex = 3;
            this.Nav_Survey.Text = "Survey";
            this.Nav_Survey.UseVisualStyleBackColor = true;
            // 
            // Home_Page
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(261, 397);
            this.Controls.Add(this.Tab_Sytem);
            this.Name = "Home_Page";
            this.Text = "Android Application";
            this.Load += new System.EventHandler(this.Home_Page_Load);
            this.Tab_Sytem.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl Tab_Sytem;
        private System.Windows.Forms.TabPage Nav_Home;
        private System.Windows.Forms.TabPage Nav_Events;
        private System.Windows.Forms.TabPage Nav_Resources;
        private System.Windows.Forms.TabPage Nav_Survey;
    }
}

