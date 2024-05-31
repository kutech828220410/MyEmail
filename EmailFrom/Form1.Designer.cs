namespace EmailFrom
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.myEmail_Send_UI1 = new MyEmail.MyEmail_Send_UI();
            this.SuspendLayout();
            // 
            // myEmail_Send_UI1
            // 
            this.myEmail_Send_UI1.EnableSsl = true;
            this.myEmail_Send_UI1.Endcoding = MyEmail.MyEmail_Send_UI.enum_Endcoding.UTF_8;
            this.myEmail_Send_UI1.FilePath = new string[] {
        "",
        "",
        "",
        ""};
            this.myEmail_Send_UI1.Location = new System.Drawing.Point(10, 10);
            this.myEmail_Send_UI1.Margin = new System.Windows.Forms.Padding(1);
            this.myEmail_Send_UI1.Name = "myEmail_Send_UI1";
            this.myEmail_Send_UI1.Port = "587";
            this.myEmail_Send_UI1.Rtf = "{\\rtf1\\ansi\\ansicpg950\\deff0\\deflang1033\\deflangfe1028{\\fonttbl{\\f0\\fnil\\fcharset" +
    "136 \\\'b7\\\'73\\\'b2\\\'d3\\\'a9\\\'fa\\\'c5\\\'e9;}}\r\n\\viewkind4\\uc1\\pard\\lang1028\\f0\\fs24\\pa" +
    "r\r\n}\r\n";
            this.myEmail_Send_UI1.SelectedRtf = "{\\rtf1\\ansi\\ansicpg950\\deff0\\deflang1033\\deflangfe1028\\uc1 }\r\n";
            this.myEmail_Send_UI1.Size = new System.Drawing.Size(745, 668);
            this.myEmail_Send_UI1.TabIndex = 0;
            this.myEmail_Send_UI1.主旨可輸入 = true;
            this.myEmail_Send_UI1.信箱收發欄位顯示 = true;
            this.myEmail_Send_UI1.傳送按鈕顯示 = true;
            this.myEmail_Send_UI1.內容可輸入 = true;
            this.myEmail_Send_UI1.副本可輸入 = true;
            this.myEmail_Send_UI1.寄件者格式要檢查 = true;
            this.myEmail_Send_UI1.密件副本可輸入 = true;
            this.myEmail_Send_UI1.捨棄按鈕顯示 = true;
            this.myEmail_Send_UI1.收件者可輸入 = true;
            this.myEmail_Send_UI1.發件者可輸入 = true;
            this.myEmail_Send_UI1.編輯欄位顯示 = true;
            this.myEmail_Send_UI1.附加檔案顯示 = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 889);
            this.Controls.Add(this.myEmail_Send_UI1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private MyEmail.MyEmail_Send_UI myEmail_Send_UI1;
    }
}

