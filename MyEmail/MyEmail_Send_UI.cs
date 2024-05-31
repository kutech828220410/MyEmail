using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
namespace MyEmail
{
    public partial class MyEmail_Send_UI : UserControl
    {
        MyThread MyThread_SendEmail;
        //创建word
        
        //创建word文档
        _Document DocFile = null;

        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialog_RTF;
        private System.Windows.Forms.SaveFileDialog saveFileDialog_RTF;
        private readonly char char_Split_Email = ';';
        private readonly string RTF_Temp_FileName = @"//content.rtf";
        private readonly string HTML_Temp_FileName = @"//content.html";
        private readonly string Temp_FilePath = @".//Email";

        #region 隱藏屬性
        [Browsable(false)]
        public override Color BackColor
        {
            get
            {
                return base.BackColor;
            }
            set
            {
                base.BackColor = value;
            }
        }
        [Browsable(false)]
        public override Image BackgroundImage
        {
            get
            {
                return base.BackgroundImage;
            }
            set
            {
                base.BackgroundImage = value;
            }
        }
        [Browsable(false)]
        public override ImageLayout BackgroundImageLayout
        {
            get
            {
                return base.BackgroundImageLayout;
            }
            set
            {
                base.BackgroundImageLayout = value;
            }
        }
        [Browsable(false)]
        public override Cursor Cursor
        {
            get
            {
                return base.Cursor;
            }
            set
            {
                base.Cursor = value;
            }
        }
        [Browsable(false)]
        public override System.Drawing.Font Font
        {
            get
            {
                return base.Font;
            }
            set
            {
                base.Font = value;
            }
        }
        [Browsable(false)]
        public override System.Drawing.Color ForeColor
        {
            get
            {
                return base.ForeColor;
            }
            set
            {
                base.ForeColor = value;
            }
        }
        [Browsable(false)]
        public override RightToLeft RightToLeft
        {
            get
            {
                return base.RightToLeft;
            }
            set
            {
                base.RightToLeft = value;
            }
        }

        #endregion
        #region 自訂屬性
        private string _UserName = "";
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public string UserName
        {
            get
            {
                return _UserName;
            }
            set
            {
                _UserName = value;
            }
        }
        private string _Password = "";
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public string Password
        {
            get
            {
                return _Password;
            }
            set
            {
                _Password = value;
            }
        }
        private string _Host = "";
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public string Host
        {
            get
            {
                return _Host;
            }
            set
            {
                _Host = value;
            }
        }
        private string _Port = "587";
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public string Port
        {
            get
            {
                return _Port;
            }
            set
            {
                _Port = value;
            }
        }
        private bool _EnableSsl = true;
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public bool EnableSsl
        {
            get
            {
                return _EnableSsl;
            }
            set
            {
                _EnableSsl = value;
            }
        }
        public enum enum_Endcoding : int
        {
            UTF_8, BIG5
        }
        private enum_Endcoding _Endcoding = enum_Endcoding.UTF_8;
        [ReadOnly(false), Browsable(true), Category("伺服器參數"), Description(""), DefaultValue("")]
        public enum_Endcoding Endcoding
        {
            get
            {
                return this._Endcoding;
            }
            set
            {
                this.comboBox_Endcoding.SelectedIndex = (int)value;
                this._Endcoding = value;
            }
        }

        private bool _寄件者格式要檢查 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 寄件者格式要檢查
        {
            get
            {
                return _寄件者格式要檢查;
            }
            set
            {
                _寄件者格式要檢查 = value;
            }
        }
        private bool _信箱收發欄位顯示 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 信箱收發欄位顯示
        {
            get
            {
                return _信箱收發欄位顯示;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.panel_信箱收發欄位.Visible = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.panel_信箱收發欄位.Visible = value; }));
                }
                _信箱收發欄位顯示 = value;
            }
        }
        private bool _編輯欄位顯示 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 編輯欄位顯示
        {
            get
            {
                return _編輯欄位顯示;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.panel_編輯欄位.Visible = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.panel_編輯欄位.Visible = value; }));
                }
                _編輯欄位顯示 = value;
            }
        }
        private bool _傳送按鈕顯示 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 傳送按鈕顯示
        {
            get
            {
                return _傳送按鈕顯示;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.button_SendEmail.Visible = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.button_SendEmail.Visible = value; }));
                }
                _傳送按鈕顯示 = value;
            }
        }
        private bool _捨棄按鈕顯示 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 捨棄按鈕顯示
        {
            get
            {
                return _捨棄按鈕顯示;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.button_Clear.Visible = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.button_Clear.Visible = value; }));
                }
                _捨棄按鈕顯示 = value;
            }
        }
        private bool _附加檔案顯示 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 附加檔案顯示
        {
            get
            {
                return _附加檔案顯示;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.panel_附加檔案_01.Visible = value;
                    this.panel_附加檔案_02.Visible = value;
                    this.panel_附加檔案_03.Visible = value;
                    this.panel_附加檔案_04.Visible = value;
                }
                else
                {
                    this.Invoke(new Action(delegate
                    {
                        this.panel_附加檔案_01.Visible = value;
                        this.panel_附加檔案_02.Visible = value;
                        this.panel_附加檔案_03.Visible = value;
                        this.panel_附加檔案_04.Visible = value;
                    }));
                }
                _附加檔案顯示 = value;
            }
        }
        private bool _發件者可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 發件者可輸入
        {
            get
            {
                return _發件者可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.textBox_發件者.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.textBox_發件者.Enabled = value; }));
                }
                _發件者可輸入 = value;
            }
        }
        private bool _收件者可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 收件者可輸入
        {
            get
            {
                return _收件者可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.textBox_收件者.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.textBox_收件者.Enabled = value; }));
                }
                _收件者可輸入 = value;
            }
        }
        private bool _副本可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 副本可輸入
        {
            get
            {
                return _副本可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.textBox_副本.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.textBox_副本.Enabled = value; }));
                }
                _副本可輸入 = value;
            }
        }
        private bool _密件副本可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 密件副本可輸入
        {
            get
            {
                return _密件副本可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.textBox_密件副本.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.textBox_密件副本.Enabled = value; }));
                }
                _密件副本可輸入 = value;
            }
        }
        private bool _主旨可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 主旨可輸入
        {
            get
            {
                return _主旨可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.textBox_主旨.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.textBox_主旨.Enabled = value; }));
                }
                _主旨可輸入 = value;
            }
        }
        private bool _內容可輸入 = true;
        [ReadOnly(false), Browsable(true), Category("介面參數"), Description(""), DefaultValue("")]
        public bool 內容可輸入
        {
            get
            {
                return _內容可輸入;
            }
            set
            {
                if (!this.IsHandleCreated)
                {
                    this.richTextBox_Email_Content.Enabled = value;
                }
                else
                {
                    this.Invoke(new Action(delegate { this.richTextBox_Email_Content.Enabled = value; }));
                }
                _內容可輸入 = value;
            }
        }
        #endregion
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Adress_From
        {
            get
            {
                return textBox_發件者.Text;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { textBox_發件者.Text = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Adress_To
        {
            get
            {
                return textBox_收件者.Text;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { textBox_收件者.Text = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Adress_CC
        {
            get
            {
                return textBox_副本.Text;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { textBox_副本.Text = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Adress_BCC
        {
            get
            {
                return textBox_密件副本.Text;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { textBox_密件副本.Text = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Subject
        {
            get
            {
                return textBox_主旨.Text;
            }
            set
            {
                this.Invoke(new Action(delegate { textBox_主旨.Text = value; }));               
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string[] FilePath
        {
            get
            {
                string[] _FilePath = new string[4];
                _FilePath[0] = textBox_附加檔案_01.Text;
                _FilePath[1] = textBox_附加檔案_02.Text;
                _FilePath[2] = textBox_附加檔案_03.Text;
                _FilePath[3] = textBox_附加檔案_04.Text;
                return _FilePath;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate
                    {
                        textBox_附加檔案_01.Text = value[0];
                        textBox_附加檔案_02.Text = value[1];
                        textBox_附加檔案_03.Text = value[2];
                        textBox_附加檔案_04.Text = value[3];
                    }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Body
        {
            get
            {
                return richTextBox_Email_Content.Text;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { richTextBox_Email_Content.Text = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Rtf
        {
            get
            {
                return richTextBox_Email_Content.Rtf;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { richTextBox_Email_Content.Rtf = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string SelectedText
        {
            get
            {
                return richTextBox_Email_Content.SelectedText;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { richTextBox_Email_Content.SelectedText = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string SelectedRtf
        {
            get
            {
                return richTextBox_Email_Content.SelectedRtf;
            }
            set
            {
                if (this.IsHandleCreated)
                {
                    this.Invoke(new Action(delegate { richTextBox_Email_Content.SelectedRtf = value; }));
                }
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public string Sender
        {
            get
            {
                if (this.savePropertyFile == null) return "";
                return savePropertyFile.Sender;
            }
            set
            {
                if (this.savePropertyFile == null) return;
                savePropertyFile.Sender = value;
            }
        }
        [ReadOnly(false), Browsable(false), Category(""), Description(""), DefaultValue("")]
        public bool IsEdited
        {
            get
            {
                bool flag = false;
                this.Invoke(new Action(delegate 
                {
                    if (this.textBox_發件者.Focused) flag = true;
                    if (this.textBox_收件者.Focused) flag = true;
                    if (this.textBox_副本.Focused) flag = true;
                    if (this.textBox_密件副本.Focused) flag = true;
                    if (this.textBox_主旨.Focused) flag = true;
                    if (this.richTextBox_Email_Content.Focused) flag = true;
                }));
                return flag;
            }
        }



        private bool _flag_BOLD = false;
        private bool flag_BOLD
        {
            get
            {
                return _flag_BOLD;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    this.Set_Button_Statu(this.button_Font_BOLD, value);
                    _flag_BOLD = value;
                }));
            }
        }
        private bool _flag_Underlined = false;
        private bool flag_Underlined
        {
            get
            {
                return _flag_Underlined;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    this.Set_Button_Statu(this.button_Font_Underlined, value);
                    _flag_Underlined = value;
                }));
            }
        }
        private bool _flag_Text_align_left = false;
        private bool flag_Text_align_left
        {
            get
            {
                return _flag_Text_align_left;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    if (value)
                    {
                        this.Set_Button_Statu(this.button_Font_Text_align_left, value);
                        this.Set_Button_Statu(this.button_Font_Text_align_center, !value);
                        this.Set_Button_Statu(this.button_Font_Text_align_right, !value);
                        _flag_Text_align_left = value;
                        _flag_Text_align_center = !value;
                        _flag_Text_align_right = !value;
                    }
            
                }));
            }
        }
        private bool _flag_Text_align_center = false;
        private bool flag_Text_align_center
        {
            get
            {
                return _flag_Text_align_center;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    if (value)
                    {
                        this.Set_Button_Statu(this.button_Font_Text_align_left, !value);
                        this.Set_Button_Statu(this.button_Font_Text_align_center, value);
                        this.Set_Button_Statu(this.button_Font_Text_align_right, !value);
                        _flag_Text_align_left = !value;
                        _flag_Text_align_center = value;
                        _flag_Text_align_right = !value;
                    }
                }));
            }
        }
        private bool _flag_Text_align_right = false;
        private bool flag_Text_align_right
        {
            get
            {
                return _flag_Text_align_right;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    if (value)
                    {
                        this.Set_Button_Statu(this.button_Font_Text_align_left, !value);
                        this.Set_Button_Statu(this.button_Font_Text_align_center, !value);
                        this.Set_Button_Statu(this.button_Font_Text_align_right, value);
                        _flag_Text_align_left = !value;
                        _flag_Text_align_center = !value;
                        _flag_Text_align_right = value;
                    }
                }));
            }
        }
        private bool _flag_Text_Color = false;
        private bool flag_Text_Color
        {
            get
            {
                return _flag_Text_Color;
            }
            set
            {
                this.Invoke(new Action(delegate
                {
                    this.Set_Button_Statu(this.button_Font_Text_Color, value);
                    _flag_Text_Color = value;
                }));
            }
        }

        private bool flag_Send_Email_Ready = false;
        private bool flag_Send_Email_error = false;
        public MyEmail_Send_UI()
        {
            InitializeComponent();
        }
    
        #region Function

        public void Init()
        {
            if(this.IsHandleCreated)
            {
                this.Invoke(new Action(delegate
                {
                    this.flag_Send_Email_Ready = true;
                    this.flag_Send_Email_error = true;

                    foreach (FontFamily font in System.Drawing.FontFamily.Families)
                    {
                        comboBox_Font_字型.Items.Add(font.Name);
                    }
                    comboBox_Font_字型.Text = "新細明體";
                    comboBox_Font_大小.Text = "12";

                    this.openFileDialog = new OpenFileDialog();
                    this.flag_Text_align_left = true;
                    richTextBox_Email_Content.Font = new System.Drawing.Font("新細明體", 12);

                    this.saveFileDialog_RTF = new SaveFileDialog();
                    this.saveFileDialog_RTF.DefaultExt = "rtf";
                    this.saveFileDialog_RTF.Filter = "rtf File (*rtf)|*rtf;";

                    this.openFileDialog_RTF = new OpenFileDialog();
                    this.openFileDialog_RTF.DefaultExt = "rtf";
                    this.openFileDialog_RTF.Filter = "rtf File (*rtf)|*rtf;";


                }));
                this.MyThread_SendEmail = new MyThread();
                this.MyThread_SendEmail.AutoRun(false);
                this.MyThread_SendEmail.Add_Method(sub_Send_Email);
                this.timer.Enabled = true;
            }
         

        }
        private void Set_Button_Statu(Button button, bool value)
        {
            this.Invoke(new Action(delegate
            {
                if (value)
                {
                    button.FlatStyle = FlatStyle.Standard;
                    button.BackColor = Color.Yellow;
                }
                else
                {
                    button.FlatStyle = FlatStyle.Standard;
                    button.BackColor = Control.DefaultBackColor;
                }
            }));

        }
        public string[] Split_Email_Adress(string email)
        {
            String[] Str_array;
            if (email != null)
            {
                Str_array = email.Split(new char[1] { char_Split_Email }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < Str_array.Length; i++) 
                {
                    Str_array[i] = Str_array[i].Replace(" ", "");
                }
            }
            else
            {
                Str_array = new string[0];
            }
            return Str_array;
        }
        public bool Check_Email_Adress(string email)
        {
            string[] Array_Email = Split_Email_Adress(email);
            if (Array_Email == null || Array_Email.Length == 0) return false;
            //Email檢查格式
            Regex EmailExpression = new Regex(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$", RegexOptions.Compiled | RegexOptions.Singleline);
            for (int i = 0; i < Array_Email.Length; i++)
            {
                try
                {
                    if (string.IsNullOrEmpty(Array_Email[i]))
                    {
                        return false;
                    }
                    else
                    {
                        if (!EmailExpression.IsMatch(Array_Email[i]))
                        {
                            return false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    //log.Error(ex.Message);
                    return false;
                }
            }
            return true;
        }
        public void Clear_FilePath(int index)
        {
            if (index >= 0 && index <= 3)
            {
                string[] filePath = this.FilePath;
                filePath[index] = "";
                this.FilePath = filePath;
            }
           
        }
        public void Set_FilePath(int index , string value)
        {
            if (index >= 0 && index <= 3)
            {
                string[] filePath = this.FilePath;
                filePath[index] = value;
                this.FilePath = filePath;
            }

        }
        public bool Add_Adress_To(string email)
        {
            if(this.Check_Email_Adress(email))
            {
                this.Adress_To = this.Adress_To + this.char_Split_Email + email;
                return true;
            }
            else
            {
                return false;
            }
        }
        public void Clear_Adress_To()
        {
            this.Adress_To = "";
        }
        public bool Add_Adress_CC(string email)
        {
            if (this.Check_Email_Adress(email))
            {
                this.Adress_CC = Adress_CC + char_Split_Email + email;
                return true;
            }
            else
            {
                return false;
            }
        }
        public void Clear_Adress_CC()
        {
            this.Adress_CC = "";
        }
        public bool Add_Adress_BCC(string email)
        {
            if (this.Check_Email_Adress(email))
            {
                this.Adress_BCC = Adress_BCC + char_Split_Email + email;
                return true;
            }
            else
            {
                return false;
            }
        }
        public void Clear_Adress_BCC()
        {
            this.Adress_BCC = "";
        }
        public void Clear_Adress_From()
        {
            this.Adress_From = "";
        }
        public void Clear_Subject()
        {
            this.Subject = "";
        }
        public void Clear_Body()
        {
            this.Body = "";
        }
        public void Clear()
        {
            this.Clear_FilePath(0);
            this.Clear_FilePath(1);
            this.Clear_FilePath(2);
            this.Clear_FilePath(3);
            this.Clear_Adress_To();
            this.Clear_Adress_CC();
            this.Clear_Adress_BCC();
            this.Clear_Adress_From();
            this.Clear_Subject();
            this.Clear_Body();
        }

        public bool Save_To_RTF(string FileName)
        {
            try
            {
                this.Invoke(new Action(delegate
                {
                    this.Invoke(new Action(delegate { richTextBox_Email_Content.SaveFile(@FileName, RichTextBoxStreamType.RichText); }));
                }));
                return true;
            }
            catch
            {
                return false;
            }         
        }
        public bool Load_To_RTF(string FileName)
        {
            try
            {
                this.Invoke(new Action(delegate
                {
                    richTextBox_Email_Content.LoadFile(@FileName, RichTextBoxStreamType.RichText);
                }));       
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void RTF_To_HTML(object LoadFileName, object SaveFileName)
        {
            LoadFileName = Path.GetFullPath((string)LoadFileName);
            SaveFileName = Path.GetFullPath((string)SaveFileName);
            _Application WordApp = new Microsoft.Office.Interop.Word.Application();
            string tmpPath = string.Empty;
            //转换文件方法
            object unknow = Type.Missing;
            //设置打开文件为RTF格式
            //object openFormat = WdSaveFormat.wdFormatRTF;// wdFormatEncodedText;// wdFormatDocument97;//.wdFormatDocument;// .wdFormatRTF;
            //object openEncoding = Microsoft.Office.Core.MsoEncoding.msoEncodingMacSimplifiedChineseGB2312;

            this.DocFile = WordApp.Documents.Open(ref LoadFileName, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
            //设置保存文件类型为html            
            //特别注意：有的机器需要设置为 WdSaveFormat.wdFormatFilteredHTML，如果出现乱码或导出的附属文件有垃圾文件，2种方式都试试
            object saveFormat = WdSaveFormat.wdFormatHTML;
            //object saveEncoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            this.DocFile.SaveAs(ref SaveFileName, ref saveFormat, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);

            this.DocFile.Close();
            WordApp.Quit();
        }
        public string Load_HTML_File(string LoadFileName)
        {
            LoadFileName = Path.GetFullPath(LoadFileName);
            //將要取得HTML原如碼的網頁放在WebRequest.Create(@”網址” )
            System.Net.WebRequest myRequest = System.Net.WebRequest.Create(@LoadFileName);

            //Method選擇GET
            myRequest.Method = "GET";

            //取得WebRequest的回覆
            System.Net.WebResponse myResponse = myRequest.GetResponse();

            //Streamreader讀取回覆
            Encoding encode = System.Text.Encoding.GetEncoding("BIG5");
            StreamReader sr = new StreamReader(myResponse.GetResponseStream(), encode);

            //將全文轉成string
            string result = sr.ReadToEnd();

            //關掉StreamReader
            sr.Close();

            //關掉WebResponse
            myResponse.Close();

            return result;
        }
        public void SaveSubject(string FileFullName)
        {
            string DirectoryName = Path.GetDirectoryName(FileFullName);
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            IFormatter binFmt = new BinaryFormatter();
            Stream stream = null;
            try
            {
                stream = File.Open(FileFullName, FileMode.Create);
                binFmt.Serialize(stream, this.Subject);
            }
            finally
            {
                if (stream != null) stream.Close();
            }
        }

        public void LoadSubject(string FileFullName)
        {
            IFormatter binFmt = new BinaryFormatter();
            Stream stream = null;
            try
            {
                if (File.Exists(FileFullName))
                {
                    stream = File.Open(FileFullName, FileMode.Open);
                    try 
                    {
                        string str = (string)binFmt.Deserialize(stream);
                        this.Subject = str; 
                    }
                    catch 
                    { 

                    }

                }
            }
            finally
            {
                if (stream != null) stream.Close();
            }
        }


        public string Get_Selection_FontName()
        {
            return this.richTextBox_Email_Content.SelectionFont.Name;
        }
        public float Get_Selection_FontSize()
        {
            return this.richTextBox_Email_Content.SelectionFont.Size;
        }
        public FontStyle Get_Selection_FontStyle()
        {
            return this.richTextBox_Email_Content.SelectionFont.Style;
        }

        public void Set_Selection_FontName(string FontName)
        {
            this.Invoke(new Action(delegate
            {
                System.Drawing.Font font = new System.Drawing.Font(FontName, this.Get_Selection_FontSize(), this.Get_Selection_FontStyle());
                richTextBox_Email_Content.SelectionFont = font;
                font.Dispose();
            }));
  
        }
        public void Set_Selection_FontSize(float FontSize)
        {
            this.Invoke(new Action(delegate
            {
                System.Drawing.Font font = new System.Drawing.Font(this.Get_Selection_FontName(), FontSize, this.Get_Selection_FontStyle());
                richTextBox_Email_Content.SelectionFont = font;           
                font.Dispose();
            }));
        }

        public HorizontalAlignment Get_Selection_Alignment()
        {
            return richTextBox_Email_Content.SelectionAlignment;
        }
        public void Set_Selection_Alignment(HorizontalAlignment HorizontalAlignment)
        {
            richTextBox_Email_Content.SelectionAlignment = HorizontalAlignment;
        }
        public void Set_Selection_Underline(bool enable)
        {
            int emun_value = 0;
            bool flag_Underline = this.Get_Selection_Underline();
            bool flag_BOLD = this.Get_Selection_BOLD();
            if (flag_BOLD)
            {
                emun_value += 1;
            }
            if (enable)
            {
                emun_value += 4;
            }
            System.Drawing.Font font = new System.Drawing.Font(this.Get_Selection_FontName(), this.Get_Selection_FontSize(), (FontStyle)emun_value);
            richTextBox_Email_Content.SelectionFont = font;
            font.Dispose();                  
        }
        public bool Get_Selection_Underline()
        {
            return richTextBox_Email_Content.SelectionFont.Underline;
        }

        public void Set_Selection_BOLD(bool enable)
        {
            int emun_value = 0;
            bool flag_Underline = this.Get_Selection_Underline();
            bool flag_BOLD = this.Get_Selection_BOLD();
            if (enable)
            {
                emun_value += 1;
            }
            if (flag_Underline)
            {
                emun_value += 4;
            }
            System.Drawing.Font font = new System.Drawing.Font(this.Get_Selection_FontName(), this.Get_Selection_FontSize(), (FontStyle)emun_value);
            richTextBox_Email_Content.SelectionFont = font;
            font.Dispose();       
        }
        public bool Get_Selection_BOLD()
        {
            return richTextBox_Email_Content.SelectionFont.Bold;
        }

        public void Set_Selection_Color(Color Color)
        {
            this.richTextBox_Email_Content.SelectionColor = Color;
        }
        public Color Get_Selection_Color()
        {
            return this.richTextBox_Email_Content.SelectionColor;
        }
        public void Set_Selection_BackColor(Color Color)
        {
            this.richTextBox_Email_Content.SelectionBackColor = Color;
        }
        public Color Get_Selection_BackColor()
        {
            return this.richTextBox_Email_Content.SelectionBackColor;
        }

        private void sub_Send_Email()
        {
            this.flag_Send_Email_Ready = false;
            this.flag_Send_Email_error = false;
            if (this.Check_All_Setting_OK())
            {
                if (!Directory.Exists(Temp_FilePath))
                {
                    Directory.CreateDirectory(Temp_FilePath);
                }
                if (this.Save_To_RTF(this.Temp_FilePath + this.RTF_Temp_FileName))
                {
                    string str_html = "";
                    this.RTF_To_HTML(this.Temp_FilePath + this.RTF_Temp_FileName, this.Temp_FilePath + this.HTML_Temp_FileName);
                    str_html = this.Load_HTML_File(this.Temp_FilePath + this.HTML_Temp_FileName);
                    File.Delete(this.Temp_FilePath + this.HTML_Temp_FileName);
                    File.Delete(this.Temp_FilePath + this.RTF_Temp_FileName);


                    System.Net.Mail.MailMessage Msg = new System.Net.Mail.MailMessage();
                    Msg.From = new System.Net.Mail.MailAddress(this.Adress_From);
                    string[] Array_Adress_To = this.Split_Email_Adress(this.Adress_To);
                    for (int i = 0; i < Array_Adress_To.Length; i++)
                    {
                        Msg.To.Add(Array_Adress_To[i]);
                    }
                    string[] Array_Adress_CC = this.Split_Email_Adress(this.Adress_CC);
                    for (int i = 0; i < Array_Adress_CC.Length; i++)
                    {
                        Msg.CC.Add(Array_Adress_CC[i]);
                    }
                    string[] Array_Adress_BCC = this.Split_Email_Adress(this.Adress_BCC);
                    for (int i = 0; i < Array_Adress_BCC.Length; i++)
                    {
                        Msg.Bcc.Add(Array_Adress_BCC[i]);
                    }

                    Msg.Subject = this.Subject;

                    Msg.Body = str_html;
                    if (this._Endcoding == enum_Endcoding.UTF_8) Msg.BodyEncoding = System.Text.Encoding.UTF8;
                    if (this._Endcoding == enum_Endcoding.BIG5) Msg.BodyEncoding = System.Text.Encoding.GetEncoding("big5");
                    Msg.IsBodyHtml = true;

                    for (int i = 0; i < FilePath.Length; i++)
                    {
                        if (FilePath[i] != "")
                        {
                            System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(FilePath[i]);
                            attachment.Name = System.IO.Path.GetFileName(FilePath[i]);
                            attachment.NameEncoding = System.Text.Encoding.UTF8;
                            attachment.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
                            attachment.ContentDisposition.Inline = true;
                            attachment.ContentDisposition.DispositionType = System.Net.Mime.DispositionTypeNames.Attachment;
                            Msg.Attachments.Add(attachment);
                        }
                    }
                    int Port = 0;
                    Port = int.Parse(this.Port);
                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(this.Host, Port);
                    smtp.EnableSsl = this.EnableSsl;
                    smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtp.Credentials = new System.Net.NetworkCredential(this.UserName, this.Password);
                    try
                    {
                        smtp.Send(Msg);
                        Msg.Attachments.Dispose();
                        Msg.Dispose();
                        this.flag_Send_Email_Ready = true;
                        this.flag_Send_Email_error = false;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        this.flag_Send_Email_Ready = true;
                        this.flag_Send_Email_error = true;
                    }

                }


            }
            else
            {
                this.flag_Send_Email_Ready = true;
                this.flag_Send_Email_error = true;
            }
        }
        public void Send_Email()
        {
            if (this.flag_Send_Email_Ready) this.MyThread_SendEmail.Trigger();
        }
        public bool Check_All_Setting_OK()
        {
            bool flag_OK = false;
            List<string> List_error_msg = new List<string>();
            string str_error_msg = "";
            if (this.UserName == "") List_error_msg.Add("未輸入使用者'名稱'");
            if (this.Password == "") List_error_msg.Add("未輸入使用者'密碼'");
            if (this.Host == "") List_error_msg.Add("未輸入伺服器'位址'");
            if (this.Port == "") List_error_msg.Add("未輸入伺服器'端口號'");
            else
            {
                int temp = 0;
                if (!int.TryParse(this.Port, out temp))
                {
                    List_error_msg.Add("伺服器'端口號'為非法字元");
                }
            }

            if (this.Adress_From == "") List_error_msg.Add("未輸入'寄件者'電子郵件位址");
            else
            {
                if (this.寄件者格式要檢查)
                {
                    if (!this.Check_Email_Adress(this.Adress_From)) List_error_msg.Add("'寄件者'電子郵件位址錯誤");
                }
        
            }      
            if (this.Adress_To == "") List_error_msg.Add("未輸入'收件者'電子郵件位址");
            else
            {
                if (!this.Check_Email_Adress(this.Adress_To)) List_error_msg.Add("'收件者'電子郵件位址錯誤");
            }
            if (this.Adress_CC != "")
            {
                if (!this.Check_Email_Adress(this.Adress_CC)) List_error_msg.Add("'副本'電子郵件位址錯誤");
            }
            if (this.Adress_BCC != "")
            {
                if (!this.Check_Email_Adress(this.Adress_BCC)) List_error_msg.Add("'密件副本'電子郵件位址錯誤");
            }

            if (this.Subject == "") List_error_msg.Add("未輸入'主旨'");


            for (int i = 0; i < FilePath.Length; i++)
            {
                if (FilePath[i] != "")
                {
                    if (!File.Exists(FilePath[i])) List_error_msg.Add("附加檔案 " + i.ToString("00")+" 錯誤");
                }
            }
            for (int i = 0; i < List_error_msg.Count; i++)
            {
                str_error_msg += i.ToString("00") + ". " + List_error_msg[i] + "\n\r";
            }
            if (str_error_msg == "") flag_OK = true;
            if (!flag_OK) MessageBox.Show(str_error_msg);
            return flag_OK;
        }

        public bool Get_Send_Ready()
        {
            return this.flag_Send_Email_Ready;
        }
        public bool Get_Send_Error()
        {
            return this.flag_Send_Email_error;
        }

        public void Replace(string old_str, string new_str)
        {
            this.richTextBox_Email_Content.Find(old_str);
            if (this.SelectedText != "")
            {
                this.SelectedText = new_str;
            }
        }
        public void Replace_RTF(string old_str, string new_RTF)
        {
            this.richTextBox_Email_Content.Find(old_str);
            if (this.SelectedRtf != "")
            {
                this.SelectedRtf = new_RTF;
            }
        }
        public class Table_Rtf
        {
            public int leftSpace = 10;        //文本与左边框的距离
            public int rowNumber = 3;        //行数
            public int colNumber = 3;        //列数
            public int tableWidth = 2000;      //单元格宽度
            public int tableHeight = 20;    //行高
            private string alignType = "center";     //居左，居中，居右
            public int redColor = 0;     //颜色
            public int greenColor = 0;
            public int blueColor = 0;
            public List<string[]> List_Table_Value = new List<string[]>();
            private System.Drawing.Font Header_Font = new System.Drawing.Font("標楷體", 12, (FontStyle)(1));
            private System.Drawing.Font Default_Font = new System.Drawing.Font("標楷體", 12);
            private System.Drawing.Color Headet_BackgroundColor = Color.White;
            private int[] ColunmWidth;
            public Table_Rtf(int colNumber, int rowNumber)
            {
                this.colNumber = colNumber;
                this.rowNumber = rowNumber;
                this.ColunmWidth = new int[colNumber];
                for (int i = 0; i < this.ColunmWidth.Length; i++)
                {
                    this.ColunmWidth[i] = 100;
                }
            }
            public void AddRow(string[] row_value)
            {
                this.List_Table_Value.Add(row_value);
            }
            public void Set_Header_Font(string Name, int Size , bool BOLD)
            {
                this.Header_Font = new System.Drawing.Font(Name, Size, (FontStyle)(BOLD ? 1 : 0));
            }
            public void Set_ColunmWidth(int Columnindex , int Width)
            {
                if(Columnindex < this.ColunmWidth.Length)
                {
                    this.ColunmWidth[Columnindex] = Width;
                }
            }
            public void Set_Header_BackgroundColor(Color color)
            {
                this.Headet_BackgroundColor = color;
            }
            public void Set_Default_Font(string Name, int Size, bool BOLD)
            {
                this.Default_Font = new System.Drawing.Font(Name, Size, (FontStyle)(BOLD ? 1 : 0));
            }
            public string Get_Table_RTF()
            {
                string RTF = "";
                RichTextBox RichTextBox = new System.Windows.Forms.RichTextBox();
                int firstWidth = this.ColunmWidth[0] - leftSpace; //第一个单元格参数，以后每个加tableWidth
                string tableStr = "{\\rtf1\\ansi\\ansicpg936\\deff0\\deflang1033\\deflangfe2052{\\fonttbl{\\f0\\fnil\\fprq2\\fcharset134 \\'cb\\'ce\\'cc\\'e5;}{\\f1\\fnil\\fcharset134\\'cb\\'ce\\'cc\\'e5;}}{\\colortbl;\\red" + redColor.ToString() + "\\green" + greenColor.ToString() + "\\blue" + blueColor.ToString() + ";}\\viewkind4\\uc1\\trowd\\trgaph" + leftSpace.ToString() + "\\trleft-" + leftSpace.ToString() + "\\trq" + alignType + "\\trbrdrt\\brdrs\\brdrw10\\brdrcf1\\trbrdrl\\brdrs\\brdrw10\\brdrcf1\\trbrdrb\\brdrs\\brdrw10\\brdrcf1\\trbrdrr\\brdrs\\brdrw10\\brdrcf1\\clbrdrt\\brdrw15\\brdrs\\clbrdrl\\brdrw15\\brdrs\\clbrdrb\\brdrw15\\brdrs\\clbrdrr\\brdrw15\\brdrs";
                int row_index = 0;
                tableStr += "\\cellx" + firstWidth.ToString() + "\\clbrdrt\\brdrw15\\brdrs\\clbrdrl\\brdrw15\\brdrs\\clbrdrb\\brdrw15\\brdrs\\clbrdrr\\brdrw15\\brdrs ";

                for (int i = 1; i < colNumber; i++)
                {
                    firstWidth += this.ColunmWidth[i];
                    tableStr += "\\cellx" + firstWidth.ToString() + "\\clbrdrt\\brdrw15\\brdrs\\clbrdrl\\brdrw15\\brdrs\\clbrdrb\\brdrw15\\brdrs\\clbrdrr\\brdrw15\\brdrs ";
                }

                //firstWidth += this.ColunmWidth[0];
                tableStr += "\\cellx" + firstWidth.ToString() + "\\pard\\intbl\\kerning2\\f0\\fs" + tableHeight.ToString();



                string cellStr = "";
                string rowStr = "";

                for (int i = 0; i < rowNumber; i++)
                {
                    cellStr = "";
                    for (int k = 0; k < colNumber; k++)
                    {
                        cellStr += @"\ansi " + "[" + (i).ToString("0") + "-" + (k).ToString("0") + "]";
                        cellStr += "\\cell";
                        row_index++;
                    }
                    rowStr = "\\intbl" + cellStr + "\\row ";
                    tableStr += rowStr;
                }

                tableStr += "\\pard\\lang2052\\kerning0\\f1\\fs18\\par }";

                RichTextBox.Rtf = tableStr;
                for (int i = 0; i < rowNumber; i++)
                {
                    for (int k = 0; k < colNumber; k++)
                    {
                        RichTextBox.SelectAll();
                        RichTextBox.Find(this.Get_RowFindStr(i, k));
                        if (i == 0)
                        {
                            RichTextBox.SelectionFont = this.Header_Font;
                            RichTextBox.SelectionBackColor = this.Headet_BackgroundColor;
                        }
                        else RichTextBox.SelectionFont = this.Default_Font;
                        if (RichTextBox.SelectedText != "")
                        {
                            if (i < List_Table_Value.Count)
                            {
                                if (k < List_Table_Value[i].Length) RichTextBox.SelectedText = List_Table_Value[i][k];
                            }
                       
                        }
                    }
                }
                RTF = RichTextBox.Rtf;
                RichTextBox.Dispose();
                return RTF;
            }
            private string Get_RowFindStr(int RowNum , int ColunmNum)
            {
                return "[" + RowNum.ToString() + "-" + ColunmNum.ToString() + "]";
            }
        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            if(this.saveFileDialog_RTF.ShowDialog(this)== System.Windows.Forms.DialogResult.OK)
            {
                this.Save_To_RTF(this.saveFileDialog_RTF.FileName);
            }
        }
        private void button_Load_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog_RTF.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                this.Load_To_RTF(this.openFileDialog_RTF.FileName);
            }
        }
        #endregion
        #region Event
        private void MyEmailUI_Load(object sender, EventArgs e)
        {
            this.Init();
        }
        private void button_Font_BOLD_Click(object sender, EventArgs e)
        {
            this.flag_BOLD = !this.flag_BOLD;
            this.Set_Selection_BOLD(this.flag_BOLD);
            button_Focus.Focus();
        }
        private void button_Font_Underlined_Click(object sender, EventArgs e)
        {
            this.flag_Underlined = !this.flag_Underlined;
            this.Set_Selection_Underline(this.flag_Underlined);
            button_Focus.Focus();
        }
        private void button_Font_Text_align_left_Click(object sender, EventArgs e)
        {
            this.flag_Text_align_left = true;
            this.Set_Selection_Alignment(HorizontalAlignment.Left);          
            button_Focus.Focus();
        }
        private void button_Font_Text_align_center_Click(object sender, EventArgs e)
        {
            this.flag_Text_align_center = true;
            this.Set_Selection_Alignment(HorizontalAlignment.Center);
            button_Focus.Focus();
        }
        private void button_Font_Text_align_right_Click(object sender, EventArgs e)
        {
            this.flag_Text_align_right = true;
            this.Set_Selection_Alignment(HorizontalAlignment.Right);
            button_Focus.Focus();
        }
        private void button_Font_Text_Color_Click(object sender, EventArgs e)
        {
            if(this.colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_Selection_Color(this.colorDialog.Color);
            }
            button_Focus.Focus();
        }
        private void button_Font_Text_BackColor_Click(object sender, EventArgs e)
        {
            if (this.BackcolorDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_Selection_BackColor(this.BackcolorDialog.Color);
            }
            button_Focus.Focus();
        }
        private void button_Focus_Click(object sender, EventArgs e)
        {
            Adress_From = textBox_發件者.Text;
        }

        private void button_附加檔案_01_清除_Click(object sender, EventArgs e)
        {
            this.Clear_FilePath(0);
        }
        private void button_附加檔案_02_清除_Click(object sender, EventArgs e)
        {
            this.Clear_FilePath(1);
        }
        private void button_附加檔案_03_清除_Click(object sender, EventArgs e)
        {
            this.Clear_FilePath(2);
        }
        private void button_附加檔案_04_清除_Click(object sender, EventArgs e)
        {
            this.Clear_FilePath(3);
        }
        private void button_附加檔案_01_瀏覽_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_FilePath(0, openFileDialog.FileName);
            }
        }
        private void button_附加檔案_02_瀏覽_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_FilePath(1, openFileDialog.FileName);
            }
        }
        private void button_附加檔案_03_瀏覽_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_FilePath(2, openFileDialog.FileName);
            }
        }
        private void button_附加檔案_04_瀏覽_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                this.Set_FilePath(3, openFileDialog.FileName);
            }
        }


        private void comboBox_Font_字型_DropDownClosed(object sender, EventArgs e)
        {
            this.Set_Selection_FontName(this.comboBox_Font_字型.Text);
        }
        private void comboBox_Font_大小_DropDownClosed(object sender, EventArgs e)
        {
            float temp;
            if (float.TryParse(comboBox_Font_大小.Text, out temp))
            {
                this.Set_Selection_FontSize(temp);
            }
        }

        private void richTextBox_Email_Content_Click(object sender, EventArgs e)
        {
            try
            {
                this.comboBox_Font_字型.Text = this.richTextBox_Email_Content.SelectionFont.Name;
                this.comboBox_Font_大小.Text = this.richTextBox_Email_Content.SelectionFont.Size.ToString();
                if (this.richTextBox_Email_Content.SelectionAlignment == HorizontalAlignment.Left) this.flag_Text_align_left = true;
                else if (this.richTextBox_Email_Content.SelectionAlignment == HorizontalAlignment.Center) this.flag_Text_align_center = true;
                else if (this.richTextBox_Email_Content.SelectionAlignment == HorizontalAlignment.Right) this.flag_Text_align_right = true;

                this.flag_BOLD = this.richTextBox_Email_Content.SelectionFont.Bold;
                this.flag_Underlined = this.richTextBox_Email_Content.SelectionFont.Underline;

                this.BackcolorDialog.Color = this.richTextBox_Email_Content.SelectionBackColor;
                this.colorDialog.Color = this.richTextBox_Email_Content.SelectionColor;
            }
            catch
            {

            }


        }
        private void button_SendEmail_Click(object sender, EventArgs e)
        {
            this.Send_Email();
        }
        private void button_Clear_Click(object sender, EventArgs e)
        {
            this.Clear_Adress_BCC();
            this.Clear_Adress_CC();
            this.Clear_Adress_From();
            this.Clear_Adress_To();
            this.Clear_FilePath(0);
            this.Clear_FilePath(1);
            this.Clear_FilePath(2);
            this.Clear_FilePath(3);
            this.Clear_Subject();
            this.richTextBox_Email_Content.Text = "";
        }
        #endregion   
        #region StreamIO
        [Serializable]
        private class SavePropertyFile
        {
            public string UserName = "";
            public string Password = "";
            public string Host = "";
            public string Port = "";
            public string Sender = "";
        }
        private SavePropertyFile savePropertyFile = new SavePropertyFile();
        public void SaveProperties()
        {
            this.SaveProperties(".\\" + this.Name + ".pro");
        }
        public void SaveProperties(string FileFullName)
        {
            string DirectoryName = Path.GetDirectoryName(FileFullName);
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            IFormatter binFmt = new BinaryFormatter();
            Stream stream = null;
            savePropertyFile.UserName = this.UserName;
            savePropertyFile.Password = this.Password;
            savePropertyFile.Host = this.Host;
            savePropertyFile.Port = this.Port;
            try
            {
                stream = File.Open(FileFullName, FileMode.Create);
                binFmt.Serialize(stream, savePropertyFile);
            }
            finally
            {
                if (stream != null) stream.Close();
            }
        }
        public void LoadProperties()
        {
            this.LoadProperties(".\\" + this.Name + ".pro");
        }
        public void LoadProperties(string FileFullName)
        {
            IFormatter binFmt = new BinaryFormatter();
            Stream stream = null;
            try
            {
                if (File.Exists(FileFullName))
                {
                    stream = File.Open(FileFullName, FileMode.Open);
                    try { savePropertyFile = (SavePropertyFile)binFmt.Deserialize(stream); }
                    catch { }

                }
                this.UserName = savePropertyFile.UserName;
                this.Password = savePropertyFile.Password;
                this.Host = savePropertyFile.Host;
                this.Port = savePropertyFile.Port;
            }
            finally
            {
                if (stream != null) stream.Close();
            }
        }
      
        #endregion
        private void timer_Tick(object sender, EventArgs e)
        {
            if (flag_Send_Email_Ready) this.label_Send_Statu.Text = "閒置中";
            else this.label_Send_Statu.Text = "發送中";
        }     
    }
    public class MyThread
    {
        public delegate void MethodDelegate();
        private bool FLAG_AutoRun = false;
        private bool FLAG_Stop = false;
        private String ThreadName = "";
        private List<MethodDelegate> Method = new List<MethodDelegate>();
        private ManualResetEvent ThreadDeadEvent, ThreadTriggerEvent;
        private System.Threading.Thread WorkerThread;
        private int SleepTime = 1;
        private double CycleTime;
        private double CycleTime_start;
        private Stopwatch stopwatch = new Stopwatch();
        private double RefreshTimeNow = 0;
        public bool IsBackGround
        {
            get
            {
                return WorkerThread.IsBackground;
            }
            set
            {
                WorkerThread.IsBackground = value;
            }
        }
        public MyThread(string ThreadName)
        {
            init(ThreadName);
        }
        public MyThread()
        {
            init("");
        }
        public MyThread(Form form)
        {
            form.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormClosing);
            init("");
        }
        public MyThread(string ThreadName, Form form)
        {
            form.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormClosing);
            init(ThreadName);
        }
        void init(string ThreadName)
        {
            // Create worker thread #1 releated ojects
            stopwatch.Start();
            this.ThreadName = ThreadName;
            ThreadDeadEvent = new ManualResetEvent(false);
            ThreadTriggerEvent = new ManualResetEvent(false);
            WorkerThread = new System.Threading.Thread(this.ThreadFunction);
            WorkerThread.IsBackground = true;
            WorkerThread.Start();
        }
        public void AutoRun(bool Enable)
        {
            FLAG_AutoRun = Enable;
        }
        public void Trigger()
        {
            ThreadTriggerEvent.Set();
        }
        public void Restart()
        {
            if (ThreadDeadEvent == null) ThreadDeadEvent = new ManualResetEvent(false);
            if (ThreadTriggerEvent == null) ThreadTriggerEvent = new ManualResetEvent(false);
            WorkerThread = new System.Threading.Thread(this.ThreadFunction);
            WorkerThread.Start();
        }
        public void Stop()
        {
            FLAG_AutoRun = false;
            FLAG_Stop = true;
            ThreadDeadEvent.Set();
            ThreadTriggerEvent.Set();
        }

        public void Add_Method(MethodDelegate method)
        {
            lock (this) Method.Add(method);
        }
        public void SetSleepTime(int Time)
        {
            this.SleepTime = Time;
        }
        public double GetCycleTime()
        {
            return Math.Round(CycleTime, 3);
        }
        public void GetCycleTime(double RefreshTime_ms, Label label)
        {
            if ((stopwatch.Elapsed.TotalMilliseconds - RefreshTimeNow) > RefreshTime_ms)
            {
                RefreshTimeNow = stopwatch.Elapsed.TotalMilliseconds;
                label.BeginInvoke(new Action(delegate
                {
                    label.Text = this.GetCycleTime().ToString();
                }));
            }
        }
        private void ThreadFunction()
        {
            MethodDelegate[] DelegateArrayUI;
            while (!ThreadDeadEvent.WaitOne(0))
            {
                if (FLAG_AutoRun) ThreadTriggerEvent.Set();
                if (FLAG_AutoRun) CycleTime_start = stopwatch.Elapsed.TotalMilliseconds;
                ThreadTriggerEvent.WaitOne();
                if (!FLAG_AutoRun) CycleTime_start = stopwatch.Elapsed.TotalMilliseconds;
                DelegateArrayUI = Method.ToArray();
                if (!FLAG_Stop)
                {
                    for (int i = 0; i < DelegateArrayUI.Length; i++)
                    {
                        if (DelegateArrayUI[i] != null) DelegateArrayUI[i]();
                    }
                }

                if (SleepTime > 0) System.Threading.Thread.Sleep(SleepTime);
                ThreadTriggerEvent.Reset();
                CycleTime = stopwatch.Elapsed.TotalMilliseconds - CycleTime_start;
            }
        }
        private void FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Stop();
        }

    }
 
}
