namespace ExcelAutomation
{
    partial class ExcelAutomation
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
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.selectedCellValue = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.selectedCell = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.displayFormula = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.selectedCellResult = new System.Windows.Forms.TextBox();
            this.submitFormula = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.selectedCellFormula = new System.Windows.Forms.TextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.destinationCell = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.sourceCell = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(324, 306);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(130, 59);
            this.button2.TabIndex = 1;
            this.button2.Text = "Submit!";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.update_cell);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(635, 299);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(152, 101);
            this.button1.TabIndex = 0;
            this.button1.Text = "Launch Excel";
            this.button1.UseMnemonic = false;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.launch_excel);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(-3, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(803, 447);
            this.tabControl1.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Location = new System.Drawing.Point(4, 34);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(795, 409);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Home";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.selectedCellValue);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.selectedCell);
            this.tabPage2.Controls.Add(this.button2);
            this.tabPage2.Location = new System.Drawing.Point(4, 34);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(795, 409);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Inputs";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(243, 178);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Validation: Numbers Only";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(231, 153);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(184, 25);
            this.label4.TabIndex = 6;
            this.label4.Text = "Selected Cell Value";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // selectedCellValue
            // 
            this.selectedCellValue.Location = new System.Drawing.Point(236, 194);
            this.selectedCellValue.Name = "selectedCellValue";
            this.selectedCellValue.Size = new System.Drawing.Size(272, 30);
            this.selectedCellValue.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(243, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Example: A1 / A1 - H1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(231, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "Select Cell";
            // 
            // selectedCell
            // 
            this.selectedCell.Location = new System.Drawing.Point(236, 93);
            this.selectedCell.Name = "selectedCell";
            this.selectedCell.Size = new System.Drawing.Size(272, 30);
            this.selectedCell.TabIndex = 2;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.displayFormula);
            this.tabPage3.Controls.Add(this.label7);
            this.tabPage3.Controls.Add(this.label8);
            this.tabPage3.Controls.Add(this.selectedCellResult);
            this.tabPage3.Controls.Add(this.submitFormula);
            this.tabPage3.Controls.Add(this.label5);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.selectedCellFormula);
            this.tabPage3.Location = new System.Drawing.Point(4, 34);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(795, 409);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Formulas";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // displayFormula
            // 
            this.displayFormula.AutoSize = true;
            this.displayFormula.Location = new System.Drawing.Point(269, 243);
            this.displayFormula.Name = "displayFormula";
            this.displayFormula.Size = new System.Drawing.Size(89, 25);
            this.displayFormula.TabIndex = 18;
            this.displayFormula.Text = "Formula:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(271, 104);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(110, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = "Example: A1 / A1 - H1";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(259, 79);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(106, 25);
            this.label8.TabIndex = 16;
            this.label8.Text = "Select Cell";
            // 
            // selectedCellResult
            // 
            this.selectedCellResult.Location = new System.Drawing.Point(264, 120);
            this.selectedCellResult.Name = "selectedCellResult";
            this.selectedCellResult.Size = new System.Drawing.Size(272, 30);
            this.selectedCellResult.TabIndex = 15;
            this.selectedCellResult.TextChanged += new System.EventHandler(this.selectedCellResult_TextChanged);
            // 
            // submitFormula
            // 
            this.submitFormula.Location = new System.Drawing.Point(341, 277);
            this.submitFormula.Name = "submitFormula";
            this.submitFormula.Size = new System.Drawing.Size(130, 59);
            this.submitFormula.TabIndex = 14;
            this.submitFormula.Text = "Submit!";
            this.submitFormula.UseVisualStyleBackColor = true;
            this.submitFormula.Click += new System.EventHandler(this.submitFormula_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(271, 194);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(160, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Validation: Mathematic operators";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(259, 169);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(131, 25);
            this.label6.TabIndex = 12;
            this.label6.Text = "Apply formula";
            // 
            // selectedCellFormula
            // 
            this.selectedCellFormula.Location = new System.Drawing.Point(264, 210);
            this.selectedCellFormula.Name = "selectedCellFormula";
            this.selectedCellFormula.Size = new System.Drawing.Size(272, 30);
            this.selectedCellFormula.TabIndex = 11;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.label12);
            this.tabPage4.Controls.Add(this.label13);
            this.tabPage4.Controls.Add(this.destinationCell);
            this.tabPage4.Controls.Add(this.label10);
            this.tabPage4.Controls.Add(this.label11);
            this.tabPage4.Controls.Add(this.sourceCell);
            this.tabPage4.Controls.Add(this.label9);
            this.tabPage4.Controls.Add(this.button8);
            this.tabPage4.Controls.Add(this.button7);
            this.tabPage4.Controls.Add(this.button6);
            this.tabPage4.Controls.Add(this.button5);
            this.tabPage4.Controls.Add(this.button4);
            this.tabPage4.Controls.Add(this.button3);
            this.tabPage4.Location = new System.Drawing.Point(4, 34);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(795, 409);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Settings";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(35, 236);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(110, 13);
            this.label12.TabIndex = 23;
            this.label12.Text = "Example: A1 / A1 - H1";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(23, 211);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(148, 25);
            this.label13.TabIndex = 22;
            this.label13.Text = "Destination Cell";
            // 
            // destinationCell
            // 
            this.destinationCell.Location = new System.Drawing.Point(28, 252);
            this.destinationCell.Name = "destinationCell";
            this.destinationCell.Size = new System.Drawing.Size(272, 30);
            this.destinationCell.TabIndex = 21;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(35, 144);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(110, 13);
            this.label10.TabIndex = 20;
            this.label10.Text = "Example: A1 / A1 - H1";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(23, 119);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(114, 25);
            this.label11.TabIndex = 19;
            this.label11.Text = "Source Cell";
            this.label11.Click += new System.EventHandler(this.label11_Click);
            // 
            // sourceCell
            // 
            this.sourceCell.Location = new System.Drawing.Point(28, 160);
            this.sourceCell.Name = "sourceCell";
            this.sourceCell.Size = new System.Drawing.Size(272, 30);
            this.sourceCell.TabIndex = 18;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(539, 211);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(181, 25);
            this.label9.TabIndex = 6;
            this.label9.Text = "Rows and Columns";
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(640, 260);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(109, 52);
            this.button8.TabIndex = 5;
            this.button8.Text = "- Column";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.remove_column);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(506, 260);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(109, 52);
            this.button7.TabIndex = 4;
            this.button7.Text = "+ Column";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.add_column);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(640, 332);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(109, 52);
            this.button6.TabIndex = 3;
            this.button6.Text = "- Row";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.remove_row);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(506, 332);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(109, 52);
            this.button5.TabIndex = 2;
            this.button5.Text = "+ Row";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.add_row);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(90, 317);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(133, 67);
            this.button4.TabIndex = 1;
            this.button4.Text = "Move data";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.move_data);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(669, 6);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(118, 42);
            this.button3.TabIndex = 0;
            this.button3.Text = "Save File";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // ExcelAutomation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tabControl1);
            this.Name = "ExcelAutomation";
            this.Text = "Excel Automation";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ExcelAutomation_FormClosing);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox selectedCell;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox selectedCellValue;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox selectedCellResult;
        private System.Windows.Forms.Button submitFormula;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox selectedCellFormula;
        private System.Windows.Forms.Label displayFormula;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox destinationCell;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox sourceCell;
    }
}

