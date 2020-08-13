namespace UnionContractWF {
    partial class Form1 {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.btn_Union = new System.Windows.Forms.Button();
            this.cmb_ContractType = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // btn_Union
            // 
            this.btn_Union.Location = new System.Drawing.Point(359, 39);
            this.btn_Union.Name = "btn_Union";
            this.btn_Union.Size = new System.Drawing.Size(89, 23);
            this.btn_Union.TabIndex = 0;
            this.btn_Union.Text = "Слияние";
            this.btn_Union.UseVisualStyleBackColor = true;
            this.btn_Union.Click += new System.EventHandler(this.btn_Union_Click);
            // 
            // cmb_ContractType
            // 
            this.cmb_ContractType.FormattingEnabled = true;
            this.cmb_ContractType.Location = new System.Drawing.Point(12, 12);
            this.cmb_ContractType.Name = "cmb_ContractType";
            this.cmb_ContractType.Size = new System.Drawing.Size(436, 21);
            this.cmb_ContractType.TabIndex = 1;
            this.cmb_ContractType.SelectedIndexChanged += new System.EventHandler(this.cmb_ContractType_SelectedIndexChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 71);
            this.Controls.Add(this.cmb_ContractType);
            this.Controls.Add(this.btn_Union);
            this.Name = "Form1";
            this.Text = "NewUnionContract";
            this.Activated += new System.EventHandler(this.Form1_Activated);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_Union;
        private System.Windows.Forms.ComboBox cmb_ContractType;
    }
}

