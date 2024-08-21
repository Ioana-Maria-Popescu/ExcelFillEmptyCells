namespace FillEmptyCells
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.fillEmptyCellsButton = new System.Windows.Forms.Button();
            this.fileNameLabel = new System.Windows.Forms.Label();
            this.failsMoveAndCountButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // fillEmptyCellsButton
            // 
            this.fillEmptyCellsButton.Location = new System.Drawing.Point(713, 49);
            this.fillEmptyCellsButton.Name = "fillEmptyCellsButton";
            this.fillEmptyCellsButton.Size = new System.Drawing.Size(75, 23);
            this.fillEmptyCellsButton.TabIndex = 0;
            this.fillEmptyCellsButton.Text = "Modifica";
            this.fillEmptyCellsButton.UseVisualStyleBackColor = true;
            this.fillEmptyCellsButton.Click += new System.EventHandler(this.fillEmptyCellsButton_Click);
            // 
            // fileNameLabel
            // 
            this.fileNameLabel.AutoSize = true;
            this.fileNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileNameLabel.Location = new System.Drawing.Point(12, 41);
            this.fileNameLabel.Name = "fileNameLabel";
            this.fileNameLabel.Size = new System.Drawing.Size(80, 31);
            this.fileNameLabel.TabIndex = 2;
            this.fileNameLabel.Text = "Label";
            // 
            // failsMoveAndCountButton
            // 
            this.failsMoveAndCountButton.Location = new System.Drawing.Point(713, 86);
            this.failsMoveAndCountButton.Name = "failsMoveAndCountButton";
            this.failsMoveAndCountButton.Size = new System.Drawing.Size(75, 23);
            this.failsMoveAndCountButton.TabIndex = 3;
            this.failsMoveAndCountButton.Text = "Modifica x2";
            this.failsMoveAndCountButton.UseVisualStyleBackColor = true;
            this.failsMoveAndCountButton.Click += new System.EventHandler(this.failsMoveAndCountButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(12, -1);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(425, 25);
            this.label2.TabIndex = 4;
            this.label2.Text = "Verificati daca prima linie din excel este corecta!";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 121);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.failsMoveAndCountButton);
            this.Controls.Add(this.fileNameLabel);
            this.Controls.Add(this.fillEmptyCellsButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "Fill Empty Cells App";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CloseApp);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button fillEmptyCellsButton;
        private System.Windows.Forms.Label fileNameLabel;
        private System.Windows.Forms.Button failsMoveAndCountButton;
        private System.Windows.Forms.Label label2;
    }
}

