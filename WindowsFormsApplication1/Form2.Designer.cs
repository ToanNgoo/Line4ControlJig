namespace WindowsFormsApplication1
{
    partial class Form2
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
            this.components = new System.ComponentModel.Container();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.label49 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lb_Jig = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox13.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.label49);
            this.groupBox13.Controls.Add(this.progressBar1);
            this.groupBox13.Controls.Add(this.lb_Jig);
            this.groupBox13.Controls.Add(this.label20);
            this.groupBox13.Location = new System.Drawing.Point(1, 3);
            this.groupBox13.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox13.Size = new System.Drawing.Size(174, 94);
            this.groupBox13.TabIndex = 20;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "Jig";
            // 
            // label49
            // 
            this.label49.AutoSize = true;
            this.label49.Location = new System.Drawing.Point(103, 18);
            this.label49.Name = "label49";
            this.label49.Size = new System.Drawing.Size(64, 13);
            this.label49.TabIndex = 27;
            this.label49.Text = "Limit: 50000";
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.progressBar1.Location = new System.Drawing.Point(6, 56);
            this.progressBar1.Maximum = 50000;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(163, 25);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 26;
            // 
            // lb_Jig
            // 
            this.lb_Jig.AutoSize = true;
            this.lb_Jig.Location = new System.Drawing.Point(41, 18);
            this.lb_Jig.Name = "lb_Jig";
            this.lb_Jig.Size = new System.Drawing.Size(13, 13);
            this.lb_Jig.TabIndex = 25;
            this.lb_Jig.Text = "0";
            this.lb_Jig.TextChanged += new System.EventHandler(this.lb_Jig_TextChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(3, 18);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(40, 13);
            this.label20.TabIndex = 20;
            this.label20.Text = "Actual:";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(177, 100);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox13);
            this.Name = "Form2";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "Form2";            
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBox13.ResumeLayout(false);
            this.groupBox13.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox13;
        private System.Windows.Forms.Label label49;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lb_Jig;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Timer timer1;
    }
}