namespace CST_With_only_excel
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.Player1 = new System.Windows.Forms.ComboBox();
            this.Player2 = new System.Windows.Forms.ComboBox();
            this.ScorePlayer1Division1 = new System.Windows.Forms.NumericUpDown();
            this.ScorePlayer2Division1 = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Report = new System.Windows.Forms.Button();
            this.ScorePlayer2Division1Game2 = new System.Windows.Forms.NumericUpDown();
            this.ScorePlayer1Division1Game2 = new System.Windows.Forms.NumericUpDown();
            this.DivisionPicker = new System.Windows.Forms.ComboBox();
            this.Finish_Button = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer1Division1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer2Division1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer2Division1Game2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer1Division1Game2)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Division:";
            // 
            // Player1
            // 
            this.Player1.FormattingEnabled = true;
            this.Player1.Location = new System.Drawing.Point(12, 70);
            this.Player1.Name = "Player1";
            this.Player1.Size = new System.Drawing.Size(121, 21);
            this.Player1.TabIndex = 1;
            this.Player1.SelectedIndexChanged += new System.EventHandler(this.Player1_SelectedIndexChanged);
            // 
            // Player2
            // 
            this.Player2.FormattingEnabled = true;
            this.Player2.Location = new System.Drawing.Point(245, 69);
            this.Player2.Name = "Player2";
            this.Player2.Size = new System.Drawing.Size(121, 21);
            this.Player2.TabIndex = 2;
            this.Player2.SelectedIndexChanged += new System.EventHandler(this.Player2_SelectedIndexChanged);
            // 
            // ScorePlayer1Division1
            // 
            this.ScorePlayer1Division1.Location = new System.Drawing.Point(139, 70);
            this.ScorePlayer1Division1.Name = "ScorePlayer1Division1";
            this.ScorePlayer1Division1.Size = new System.Drawing.Size(40, 20);
            this.ScorePlayer1Division1.TabIndex = 3;
            // 
            // ScorePlayer2Division1
            // 
            this.ScorePlayer2Division1.Location = new System.Drawing.Point(199, 70);
            this.ScorePlayer2Division1.Name = "ScorePlayer2Division1";
            this.ScorePlayer2Division1.Size = new System.Drawing.Size(40, 20);
            this.ScorePlayer2Division1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(83, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Spelare1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(316, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(49, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Spelare2";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(162, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(38, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Poäng";
            // 
            // Report
            // 
            this.Report.Location = new System.Drawing.Point(381, 67);
            this.Report.Name = "Report";
            this.Report.Size = new System.Drawing.Size(129, 23);
            this.Report.TabIndex = 48;
            this.Report.Text = "Rapportera resultat";
            this.Report.UseVisualStyleBackColor = true;
            this.Report.Click += new System.EventHandler(this.ReportDivision1_Click);
            // 
            // ScorePlayer2Division1Game2
            // 
            this.ScorePlayer2Division1Game2.Location = new System.Drawing.Point(199, 100);
            this.ScorePlayer2Division1Game2.Name = "ScorePlayer2Division1Game2";
            this.ScorePlayer2Division1Game2.Size = new System.Drawing.Size(40, 20);
            this.ScorePlayer2Division1Game2.TabIndex = 55;
            // 
            // ScorePlayer1Division1Game2
            // 
            this.ScorePlayer1Division1Game2.Location = new System.Drawing.Point(139, 100);
            this.ScorePlayer1Division1Game2.Name = "ScorePlayer1Division1Game2";
            this.ScorePlayer1Division1Game2.Size = new System.Drawing.Size(40, 20);
            this.ScorePlayer1Division1Game2.TabIndex = 54;
            // 
            // DivisionPicker
            // 
            this.DivisionPicker.FormattingEnabled = true;
            this.DivisionPicker.Location = new System.Drawing.Point(68, 1);
            this.DivisionPicker.Name = "DivisionPicker";
            this.DivisionPicker.Size = new System.Drawing.Size(121, 21);
            this.DivisionPicker.TabIndex = 56;
            this.DivisionPicker.SelectedIndexChanged += new System.EventHandler(this.DivisionPicker_SelectedIndexChanged);
            // 
            // Finish_Button
            // 
            this.Finish_Button.Location = new System.Drawing.Point(1101, 598);
            this.Finish_Button.Name = "Finish_Button";
            this.Finish_Button.Size = new System.Drawing.Size(129, 23);
            this.Finish_Button.TabIndex = 57;
            this.Finish_Button.Text = "Klart";
            this.Finish_Button.UseVisualStyleBackColor = true;
            this.Finish_Button.Click += new System.EventHandler(this.Finish_Button_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1242, 633);
            this.Controls.Add(this.Finish_Button);
            this.Controls.Add(this.DivisionPicker);
            this.Controls.Add(this.ScorePlayer2Division1Game2);
            this.Controls.Add(this.ScorePlayer1Division1Game2);
            this.Controls.Add(this.Report);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ScorePlayer2Division1);
            this.Controls.Add(this.ScorePlayer1Division1);
            this.Controls.Add(this.Player2);
            this.Controls.Add(this.Player1);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer1Division1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer2Division1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer2Division1Game2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ScorePlayer1Division1Game2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox Player1;
        private System.Windows.Forms.ComboBox Player2;
        private System.Windows.Forms.NumericUpDown ScorePlayer1Division1;
        private System.Windows.Forms.NumericUpDown ScorePlayer2Division1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button Report;
        private System.Windows.Forms.NumericUpDown ScorePlayer2Division1Game2;
        private System.Windows.Forms.NumericUpDown ScorePlayer1Division1Game2;
        private System.Windows.Forms.ComboBox DivisionPicker;
        private System.Windows.Forms.Button Finish_Button;
    }
}

