namespace ExcelExtractor
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
            this.exelBox = new System.Windows.Forms.GroupBox();
            this.lblRezult = new System.Windows.Forms.Label();
            this.btnExcel = new System.Windows.Forms.Button();
            this.exelLbl = new System.Windows.Forms.Label();
            this.wordBox = new System.Windows.Forms.GroupBox();
            this.lblFolder = new System.Windows.Forms.Label();
            this.btnConvert = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnWord = new System.Windows.Forms.Button();
            this.wordLbl = new System.Windows.Forms.Label();
            this.message = new System.Windows.Forms.Label();
            this.anulare = new System.Windows.Forms.Button();
            this.exelBox.SuspendLayout();
            this.wordBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // exelBox
            // 
            this.exelBox.Controls.Add(this.lblRezult);
            this.exelBox.Controls.Add(this.btnExcel);
            this.exelBox.Controls.Add(this.exelLbl);
            this.exelBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exelBox.Location = new System.Drawing.Point(13, 12);
            this.exelBox.Name = "exelBox";
            this.exelBox.Size = new System.Drawing.Size(290, 247);
            this.exelBox.TabIndex = 0;
            this.exelBox.TabStop = false;
            this.exelBox.Text = "Opțiuni Document Excel";
            // 
            // lblRezult
            // 
            this.lblRezult.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRezult.Location = new System.Drawing.Point(7, 120);
            this.lblRezult.Name = "lblRezult";
            this.lblRezult.Size = new System.Drawing.Size(277, 116);
            this.lblRezult.TabIndex = 2;
            // 
            // btnExcel
            // 
            this.btnExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExcel.Location = new System.Drawing.Point(10, 83);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(91, 34);
            this.btnExcel.TabIndex = 1;
            this.btnExcel.Text = "Document";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // exelLbl
            // 
            this.exelLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exelLbl.Location = new System.Drawing.Point(6, 33);
            this.exelLbl.Name = "exelLbl";
            this.exelLbl.Size = new System.Drawing.Size(278, 47);
            this.exelLbl.TabIndex = 0;
            this.exelLbl.Text = "1. Încărcați documentul Excel din care doriți extragerea textului!";
            // 
            // wordBox
            // 
            this.wordBox.Controls.Add(this.anulare);
            this.wordBox.Controls.Add(this.lblFolder);
            this.wordBox.Controls.Add(this.btnConvert);
            this.wordBox.Controls.Add(this.label1);
            this.wordBox.Controls.Add(this.btnWord);
            this.wordBox.Controls.Add(this.wordLbl);
            this.wordBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.wordBox.Location = new System.Drawing.Point(309, 12);
            this.wordBox.Name = "wordBox";
            this.wordBox.Size = new System.Drawing.Size(307, 247);
            this.wordBox.TabIndex = 1;
            this.wordBox.TabStop = false;
            this.wordBox.Text = "Opțiuni Document Word";
            // 
            // lblFolder
            // 
            this.lblFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolder.Location = new System.Drawing.Point(7, 120);
            this.lblFolder.Name = "lblFolder";
            this.lblFolder.Size = new System.Drawing.Size(294, 32);
            this.lblFolder.TabIndex = 5;
            // 
            // btnConvert
            // 
            this.btnConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConvert.Location = new System.Drawing.Point(10, 202);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(91, 34);
            this.btnConvert.TabIndex = 4;
            this.btnConvert.Text = "Copiază";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 152);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(295, 47);
            this.label1.TabIndex = 3;
            this.label1.Text = "3. Copiați conținutul fiecărei celule din Excel într-un document Word separat!";
            // 
            // btnWord
            // 
            this.btnWord.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWord.Location = new System.Drawing.Point(10, 83);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(91, 34);
            this.btnWord.TabIndex = 2;
            this.btnWord.Text = "Locație";
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // wordLbl
            // 
            this.wordLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.wordLbl.Location = new System.Drawing.Point(6, 33);
            this.wordLbl.Name = "wordLbl";
            this.wordLbl.Size = new System.Drawing.Size(295, 47);
            this.wordLbl.TabIndex = 1;
            this.wordLbl.Text = "2. Alegți locația în care doriți să fie salvate documentele word!";
            // 
            // message
            // 
            this.message.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.message.Location = new System.Drawing.Point(12, 262);
            this.message.Name = "message";
            this.message.Size = new System.Drawing.Size(604, 78);
            this.message.TabIndex = 2;
            // 
            // anulare
            // 
            this.anulare.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.anulare.Location = new System.Drawing.Point(210, 202);
            this.anulare.Name = "anulare";
            this.anulare.Size = new System.Drawing.Size(91, 34);
            this.anulare.TabIndex = 6;
            this.anulare.Text = "Anulare";
            this.anulare.UseVisualStyleBackColor = true;
            this.anulare.Click += new System.EventHandler(this.anulare_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(630, 349);
            this.Controls.Add(this.message);
            this.Controls.Add(this.wordBox);
            this.Controls.Add(this.exelBox);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelToWord";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.exelBox.ResumeLayout(false);
            this.wordBox.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox exelBox;
        private System.Windows.Forms.Label exelLbl;
        private System.Windows.Forms.GroupBox wordBox;
        private System.Windows.Forms.Label wordLbl;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Label lblRezult;
        private System.Windows.Forms.Label lblFolder;
        private System.Windows.Forms.Label message;
        private System.Windows.Forms.Button anulare;
    }
}

