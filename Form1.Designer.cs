namespace XmlCreator
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
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtMain = new System.Windows.Forms.TextBox();
            this.lbl1 = new System.Windows.Forms.Label();
            this.btnStart = new System.Windows.Forms.Button();
            this.butDest = new System.Windows.Forms.Button();
            this.txtDest = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnArchive = new System.Windows.Forms.Button();
            this.txtArchive = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblRow = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblOf = new System.Windows.Forms.Label();
            this.lblMsg = new System.Windows.Forms.Label();
            this.lblMsgPdfError = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(579, 22);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 10;
            this.btnBrowse.Text = "חפש";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtMain
            // 
            this.txtMain.Location = new System.Drawing.Point(61, 24);
            this.txtMain.Name = "txtMain";
            this.txtMain.Size = new System.Drawing.Size(495, 20);
            this.txtMain.TabIndex = 9;
            // 
            // lbl1
            // 
            this.lbl1.AutoSize = true;
            this.lbl1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl1.Location = new System.Drawing.Point(662, 26);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(79, 19);
            this.lbl1.TabIndex = 8;
            this.lbl1.Text = "קובץ אקסל";
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(617, 240);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 11;
            this.btnStart.Text = "התחל";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // butDest
            // 
            this.butDest.Location = new System.Drawing.Point(579, 71);
            this.butDest.Name = "butDest";
            this.butDest.Size = new System.Drawing.Size(75, 23);
            this.butDest.TabIndex = 14;
            this.butDest.Text = "חפש";
            this.butDest.UseVisualStyleBackColor = true;
            this.butDest.Click += new System.EventHandler(this.butDest_Click);
            // 
            // txtDest
            // 
            this.txtDest.Location = new System.Drawing.Point(61, 71);
            this.txtDest.Name = "txtDest";
            this.txtDest.Size = new System.Drawing.Size(495, 20);
            this.txtDest.TabIndex = 13;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label1.Location = new System.Drawing.Point(668, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 19);
            this.label1.TabIndex = 12;
            this.label1.Text = "תקיית יעד";
            // 
            // btnArchive
            // 
            this.btnArchive.Location = new System.Drawing.Point(579, 128);
            this.btnArchive.Name = "btnArchive";
            this.btnArchive.Size = new System.Drawing.Size(75, 23);
            this.btnArchive.TabIndex = 17;
            this.btnArchive.Text = "חפש";
            this.btnArchive.UseVisualStyleBackColor = true;
            this.btnArchive.Click += new System.EventHandler(this.btnArchive_Click);
            // 
            // txtArchive
            // 
            this.txtArchive.Location = new System.Drawing.Point(61, 128);
            this.txtArchive.Name = "txtArchive";
            this.txtArchive.Size = new System.Drawing.Size(495, 20);
            this.txtArchive.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label2.Location = new System.Drawing.Point(668, 132);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 19);
            this.label2.TabIndex = 15;
            this.label2.Text = "תקיית ארכיון";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(501, 240);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 18;
            this.button1.Text = "סגור";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label3.Location = new System.Drawing.Point(88, 193);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 16);
            this.label3.TabIndex = 19;
            this.label3.Text = "Working on row";
            // 
            // lblRow
            // 
            this.lblRow.AutoSize = true;
            this.lblRow.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblRow.Location = new System.Drawing.Point(212, 193);
            this.lblRow.Name = "lblRow";
            this.lblRow.Size = new System.Drawing.Size(15, 16);
            this.lblRow.TabIndex = 20;
            this.lblRow.Text = "x";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label5.Location = new System.Drawing.Point(247, 193);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(23, 16);
            this.label5.TabIndex = 21;
            this.label5.Text = "Of";
            // 
            // lblOf
            // 
            this.lblOf.AutoSize = true;
            this.lblOf.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblOf.Location = new System.Drawing.Point(294, 193);
            this.lblOf.Name = "lblOf";
            this.lblOf.Size = new System.Drawing.Size(15, 16);
            this.lblOf.TabIndex = 22;
            this.lblOf.Text = "x";
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblMsg.Location = new System.Drawing.Point(88, 193);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(0, 16);
            this.lblMsg.TabIndex = 23;
            // 
            // lblMsgPdfError
            // 
            this.lblMsgPdfError.AutoSize = true;
            this.lblMsgPdfError.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblMsgPdfError.Location = new System.Drawing.Point(97, 291);
            this.lblMsgPdfError.Name = "lblMsgPdfError";
            this.lblMsgPdfError.Size = new System.Drawing.Size(0, 16);
            this.lblMsgPdfError.TabIndex = 24;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lblMsgPdfError);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.lblOf);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lblRow);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnArchive);
            this.Controls.Add(this.txtArchive);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.butDest);
            this.Controls.Add(this.txtDest);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtMain);
            this.Controls.Add(this.lbl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtMain;
        private System.Windows.Forms.Label lbl1;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button butDest;
        private System.Windows.Forms.TextBox txtDest;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnArchive;
        private System.Windows.Forms.TextBox txtArchive;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblRow;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblOf;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.Label lblMsgPdfError;
    }
}

