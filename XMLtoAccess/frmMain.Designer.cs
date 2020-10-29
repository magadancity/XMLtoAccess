namespace XMLtoAccess
{
    partial class frmMain
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlFill = new System.Windows.Forms.Panel();
            this.pnlElem = new System.Windows.Forms.Panel();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnSelectPathToDb = new System.Windows.Forms.Button();
            this.txtPathToDB = new System.Windows.Forms.TextBox();
            this.lblSelectPathToDb = new System.Windows.Forms.Label();
            this.btnSelectXml = new System.Windows.Forms.Button();
            this.txtPathToArc = new System.Windows.Forms.TextBox();
            this.lblSelectPathToXML = new System.Windows.Forms.Label();
            this.pnlBottom = new System.Windows.Forms.Panel();
            this.btnRun = new System.Windows.Forms.Button();
            this.pb = new System.Windows.Forms.ProgressBar();
            this.pnlFill.SuspendLayout();
            this.pnlElem.SuspendLayout();
            this.pnlBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlFill
            // 
            this.pnlFill.Controls.Add(this.pnlElem);
            this.pnlFill.Controls.Add(this.pnlBottom);
            this.pnlFill.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFill.Location = new System.Drawing.Point(0, 0);
            this.pnlFill.Name = "pnlFill";
            this.pnlFill.Size = new System.Drawing.Size(702, 349);
            this.pnlFill.TabIndex = 0;
            // 
            // pnlElem
            // 
            this.pnlElem.Controls.Add(this.txtLog);
            this.pnlElem.Controls.Add(this.btnSelectPathToDb);
            this.pnlElem.Controls.Add(this.txtPathToDB);
            this.pnlElem.Controls.Add(this.lblSelectPathToDb);
            this.pnlElem.Controls.Add(this.btnSelectXml);
            this.pnlElem.Controls.Add(this.txtPathToArc);
            this.pnlElem.Controls.Add(this.lblSelectPathToXML);
            this.pnlElem.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlElem.Location = new System.Drawing.Point(0, 0);
            this.pnlElem.Name = "pnlElem";
            this.pnlElem.Size = new System.Drawing.Size(702, 293);
            this.pnlElem.TabIndex = 3;
            // 
            // txtLog
            // 
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Location = new System.Drawing.Point(15, 92);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.Size = new System.Drawing.Size(675, 195);
            this.txtLog.TabIndex = 6;
            // 
            // btnSelectPathToDb
            // 
            this.btnSelectPathToDb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectPathToDb.Location = new System.Drawing.Point(617, 52);
            this.btnSelectPathToDb.Name = "btnSelectPathToDb";
            this.btnSelectPathToDb.Size = new System.Drawing.Size(73, 34);
            this.btnSelectPathToDb.TabIndex = 5;
            this.btnSelectPathToDb.Text = "Выбрать";
            this.btnSelectPathToDb.UseVisualStyleBackColor = true;
            this.btnSelectPathToDb.Click += new System.EventHandler(this.btnSelectPathToDb_Click);
            // 
            // txtPathToDB
            // 
            this.txtPathToDB.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPathToDB.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPathToDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtPathToDB.Location = new System.Drawing.Point(153, 52);
            this.txtPathToDB.Multiline = true;
            this.txtPathToDB.Name = "txtPathToDB";
            this.txtPathToDB.ReadOnly = true;
            this.txtPathToDB.Size = new System.Drawing.Size(458, 34);
            this.txtPathToDB.TabIndex = 4;
            // 
            // lblSelectPathToDb
            // 
            this.lblSelectPathToDb.Location = new System.Drawing.Point(12, 52);
            this.lblSelectPathToDb.Name = "lblSelectPathToDb";
            this.lblSelectPathToDb.Size = new System.Drawing.Size(135, 34);
            this.lblSelectPathToDb.TabIndex = 3;
            this.lblSelectPathToDb.Text = "Укажите путь к БД";
            this.lblSelectPathToDb.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnSelectXml
            // 
            this.btnSelectXml.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelectXml.Location = new System.Drawing.Point(617, 9);
            this.btnSelectXml.Name = "btnSelectXml";
            this.btnSelectXml.Size = new System.Drawing.Size(73, 34);
            this.btnSelectXml.TabIndex = 2;
            this.btnSelectXml.Text = "Выбрать";
            this.btnSelectXml.UseVisualStyleBackColor = true;
            this.btnSelectXml.Click += new System.EventHandler(this.btnSelectXml_Click);
            // 
            // txtPathToArc
            // 
            this.txtPathToArc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPathToArc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPathToArc.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtPathToArc.Location = new System.Drawing.Point(153, 9);
            this.txtPathToArc.Multiline = true;
            this.txtPathToArc.Name = "txtPathToArc";
            this.txtPathToArc.ReadOnly = true;
            this.txtPathToArc.Size = new System.Drawing.Size(458, 34);
            this.txtPathToArc.TabIndex = 1;
            // 
            // lblSelectPathToXML
            // 
            this.lblSelectPathToXML.Location = new System.Drawing.Point(12, 9);
            this.lblSelectPathToXML.Name = "lblSelectPathToXML";
            this.lblSelectPathToXML.Size = new System.Drawing.Size(135, 34);
            this.lblSelectPathToXML.TabIndex = 0;
            this.lblSelectPathToXML.Text = "Укажите файл архива";
            this.lblSelectPathToXML.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pnlBottom
            // 
            this.pnlBottom.Controls.Add(this.pb);
            this.pnlBottom.Controls.Add(this.btnRun);
            this.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlBottom.Location = new System.Drawing.Point(0, 293);
            this.pnlBottom.Name = "pnlBottom";
            this.pnlBottom.Size = new System.Drawing.Size(702, 56);
            this.pnlBottom.TabIndex = 2;
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(565, 6);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(125, 38);
            this.btnRun.TabIndex = 0;
            this.btnRun.Text = "Создать базу данных";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(15, 6);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(459, 23);
            this.pb.TabIndex = 1;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(702, 349);
            this.Controls.Add(this.pnlFill);
            this.Name = "frmMain";
            this.Text = "Создание БД";
            this.pnlFill.ResumeLayout(false);
            this.pnlElem.ResumeLayout(false);
            this.pnlElem.PerformLayout();
            this.pnlBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlFill;
        private System.Windows.Forms.Panel pnlElem;
        private System.Windows.Forms.Panel pnlBottom;
        private System.Windows.Forms.Button btnSelectXml;
        private System.Windows.Forms.TextBox txtPathToArc;
        private System.Windows.Forms.Label lblSelectPathToXML;
        private System.Windows.Forms.Button btnSelectPathToDb;
        private System.Windows.Forms.TextBox txtPathToDB;
        private System.Windows.Forms.Label lblSelectPathToDb;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.ProgressBar pb;
    }
}

