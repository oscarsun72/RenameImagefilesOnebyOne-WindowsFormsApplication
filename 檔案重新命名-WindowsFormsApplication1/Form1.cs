using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;
using wfrm = System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using io = System.IO;
using static System.Net.Mime.MediaTypeNames;

namespace 檔案重新命名_WindowsFormsApplication1
{
    public partial class Form1 : wfrm.Form
    {
        Process prcssDownloadImgFullName;
        //副檔名清單
        readonly string[] imgExt = { ".jpg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".jpeg" };
        List<string> imagesList = new List<string>();
        string sourcePath = "";
        public Form1()
        {
            InitializeComponent();
            SimpleList family = new SimpleList();

            // Populate the List            
            family.Add("沛榮");
            family.Add("玫儀");
            family.Add("靖F");
            family.Add("JT");
            family.Add("five");
            family.Add("six");
            family.Add("seven");
            family.Add("eight");
            family.PrintContents();
            comboBox1.DataSource = family;
            pictureBox1.MouseWheel += new MouseEventHandler(pictureBox1_MouseWheel);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {

            ////test the pull request!
            //if (io.Directory.Exists(textBox1.Text) && io.Directory.Exists(textBox2.Text))
            //{
            //    foreach (var itemF in io.Directory.GetFiles(textBox1.Text))
            //    {
            //        io.FileInfo f = new io.FileInfo(itemF);
            //        foreach (var itemT in io.Directory.GetFiles(textBox2.Text))
            //        {
            //            io.FileInfo ft = new io.FileInfo(itemT);
            //            if (f.LastWriteTime == ft.LastWriteTime && f.Length == ft.Length && f.Name == ft.Name)
            //            {

            //                if (io.File.GetAttributes(itemF).ToString().IndexOf(io.FileAttributes.ReadOnly.ToString()) != -1)
            //                {
            //                    io.File.SetAttributes(itemF, io.FileAttributes.Normal);
            //                }
            //                io.File.Delete(itemF);
            //                break;
            //            }
            //        }
            //    }
            //    MessageBox.Show("done!");
            //}
            //else
            //{
            //    MessageBox.Show("路徑有誤！");
            //}
        }


        private void textBox1_Click(object sender, EventArgs e)
        {
            //textBox1.Text = "";            
            textBox1.Text = Clipboard.GetText();
        }
        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox2.Text = Clipboard.GetText();
        }


        private void Form1_DragEnter(object sender, DragEventArgs e)
        {//https://wijtb.nctu.me/archives/269/
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;//调用DragDrop事件 
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {//https://www.google.com/search?rlz=1C1OKWM_zh-TWTW847TW847&ei=W7wBXbDjG8Kk8AWfkYGoBQ&q=c%23+%E6%8B%96%E6%8B%89%E7%89%A9%E4%BB%B6+%E8%B3%87%E6%96%99%E5%A4%BE&oq=c%23+%E6%8B%96%E6%8B%89%E7%89%A9%E4%BB%B6+%E8%B3%87%E6%96%99&gs_l=psy-ab.3.0.33i160l2.11461.20026..22176...0.0..0.99.232.3......0....1..gws-wiz.......0i30.KqdNNbPjI8A
            //https://wijtb.nctu.me/archives/269/
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.textBox1.Text = filePaths[0];
        }

        private void textBox2_DragDrop(object sender, DragEventArgs e)
        {
            string[] filePaths = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.textBox2.Text = filePaths[0];
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (MessageBox.Show("結束應用程式？", "確定結束？", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    == DialogResult.OK)
                    this.Close();
            }

            //瀏覽圖檔
            switch (e.KeyCode)
            {
                case Keys.Home:
                    //第一張圖檔
                    if (imagesList.Count() == 0) return;
                    imageIndex = 0;
                    pictureBox1.ImageLocation = imagesList[imageIndex];
                    //顯示圖檔全檔名
                    textBox1.Text = pictureBox1.ImageLocation;
                    break;
                case Keys.End:
                    //最後一張圖檔
                    if (imagesList.Count() == 0) return;
                    imageIndex = imagesList.Count - 1;
                    pictureBox1.ImageLocation = imagesList[imageIndex];
                    //顯示圖檔全檔名
                    textBox1.Text = pictureBox1.ImageLocation;
                    break;
                case Keys.Down:
                    nextImageShow();
                    break;
                case Keys.Right:
                    nextImageShow();
                    break;
                case Keys.PageDown:
                    nextImageShow();
                    break;
                case Keys.Up:
                    previousImageShow();
                    break;
                case Keys.Left:
                    previousImageShow();
                    break;
                case Keys.PageUp:
                    previousImageShow();
                    break;
                default:
                    break;
            }
        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void comboBox_Click(wfrm.ComboBox cmbx)
        {
            string x = cmbx.SelectedValue.ToString();
            if (textBox2.Text == "重新命名預覽")
            {
                textBox2.Text = "";
            }
            textBox2.Text = textBox2.Text + x;

        }

        class SimpleList : IList //https://docs.microsoft.com/zh-tw/dotnet/api/system.collections.ilist?view=netframework-4.8
        {
            private object[] _contents = new object[8];
            private int _count;

            public SimpleList()
            {
                _count = 0;
            }

            // IList Members
            public int Add(object value)
            {
                if (_count < _contents.Length)
                {
                    _contents[_count] = value;
                    _count++;

                    return (_count - 1);
                }
                else
                {
                    return -1;
                }
            }

            public void Clear()
            {
                _count = 0;
            }

            public bool Contains(object value)
            {
                bool inList = false;
                for (int i = 0; i < Count; i++)
                {
                    if (_contents[i] == value)
                    {
                        inList = true;
                        break;
                    }
                }
                return inList;
            }

            public int IndexOf(object value)
            {
                int itemIndex = -1;
                for (int i = 0; i < Count; i++)
                {
                    if (_contents[i] == value)
                    {
                        itemIndex = i;
                        break;
                    }
                }
                return itemIndex;
            }

            public void Insert(int index, object value)
            {
                if ((_count + 1 <= _contents.Length) && (index < Count) && (index >= 0))
                {
                    _count++;

                    for (int i = Count - 1; i > index; i--)
                    {
                        _contents[i] = _contents[i - 1];
                    }
                    _contents[index] = value;
                }
            }

            public bool IsFixedSize
            {
                get
                {
                    return true;
                }
            }

            public bool IsReadOnly
            {
                get
                {
                    return false;
                }
            }

            public void Remove(object value)
            {
                RemoveAt(IndexOf(value));
            }

            public void RemoveAt(int index)
            {
                if ((index >= 0) && (index < Count))
                {
                    for (int i = index; i < Count - 1; i++)
                    {
                        _contents[i] = _contents[i + 1];
                    }
                    _count--;
                }
            }

            public object this[int index]
            {
                get
                {
                    return _contents[index];
                }
                set
                {
                    _contents[index] = value;
                }
            }

            // ICollection Members

            public void CopyTo(Array array, int index)
            {
                int j = index;
                for (int i = 0; i < Count; i++)
                {
                    array.SetValue(_contents[i], j);
                    j++;
                }
            }

            public int Count
            {
                get
                {
                    return _count;
                }
            }

            public bool IsSynchronized
            {
                get
                {
                    return false;
                }
            }

            // Return the current instance since the underlying store is not
            // publicly available.
            public object SyncRoot
            {
                get
                {
                    return this;
                }
            }

            // IEnumerable Members

            public IEnumerator GetEnumerator()
            {
                // Refer to the IEnumerator documentation for an example of
                // implementing an enumerator.
                throw new Exception("The method or operation is not implemented.");
            }

            public void PrintContents()
            {
                Console.WriteLine("List has a capacity of {0} and currently has {1} elements.", _contents.Length, _count);
                Console.Write("List contents:");
                for (int i = 0; i < Count; i++)
                {
                    Console.Write(" {0}", _contents[i]);
                }
                Console.WriteLine();
            }
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            //comboBox1.DroppedDown=true;

        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBox_Click(comboBox1);
        }

        private void comboBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Middle)
            {
                comboBox_Click(comboBox1);

            }

        }

        private void comboBox1_Click_1(object sender, EventArgs e)
        {
            comboBox1.DroppedDown = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                clearImageList();
                return;
            }
            //if (MessageBox.Show("是否開始資料夾內的檔案重新命名？", "重新命名", MessageBoxButtons.OKCancel) == DialogResult.Cancel) return;           
            string path = textBox1.Text;
            //處理資料夾內檔案
            if (io.Directory.Exists(path))
            {
                if (!loadImageList(path)) return;
                //若有圖檔則載入第一張圖檔
                imageIndex = 0;
                pictureBox1.ImageLocation = imagesList[imageIndex];
                textBox1.Text = pictureBox1.ImageLocation;


                //foreach (string fstr in filesList) //io.Directory.GetFiles(path))
                //{
                //    if (fstr == "") continue;
                //    io.FileInfo f = new io.FileInfo(fstr);
                //    //if (f.LastAccessTime< DateTime.Today)
                //    //{
                //    switch (f.Extension)
                //    {
                //        case ".jpg":
                //            pictureBox1.ImageLocation = f.FullName;
                //            // pictureBox1.Image = new Bitmap( f.FullName);//二式皆可
                //            return;
                //        //break;
                //        case ".jpeg":
                //        case ".png":
                //        case ".bmp":
                //        case ".tif":
                //        case ".tiff":
                //        default:
                //            break;
                //    }
                //    //}

                //}
            }
            //處理單一檔案
            else
            {
                //如果是輸入全檔名則載入該圖檔瀏覽
                if (path != pictureBox1.ImageLocation)
                    pictureBox1.ImageLocation = path;
                //如果不是則僅限檢視目前圖檔全檔名


            }
        }

        //清除圖檔清單
        private void clearImageList()
        {
            imagesList.Clear();
            imageIndex = 0;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        int imageIndex = 0;
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.Control)
            {
                openByMsPaint();
                return;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Middle)
            {

            }
        }

        //202301110038 creedit chatGPT：C# Windows Form MouseWheelEvent：感恩感恩　讚歎讚歎　南無阿彌陀佛
        private void pictureBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta > 0)
            {
                // 滾輪向上
                previousImageShow();
            }
            else
            {
                // 滾輪向下
                nextImageShow();
            }
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Left:
                    if (ModifierKeys == Keys.None)
                    {
                        nextImageShow();
                    }
                    break;
                case MouseButtons.None:
                    break;
                case MouseButtons.Right:
                    if (ModifierKeys == Keys.None)
                    {
                        previousImageShow();
                    }
                    break;
                case MouseButtons.Middle:
                    break;
                case MouseButtons.XButton1:
                    previousImageShow();
                    break;
                case MouseButtons.XButton2:
                    nextImageShow();
                    break;
                default:
                    break;
            }

        }

        private void previousImageShow()
        {
            if (imagesList.Count() == 0) return;
            if (imageIndex < 0) { imageIndex = 0; return; }
            //next image
            while (imageIndex - 1 > -1 && !File.Exists(imagesList[--imageIndex])) { }
            pictureBox1.ImageLocation = imagesList[imageIndex];
            //顯示圖檔全檔名
            textBox1.Text = pictureBox1.ImageLocation;
            if (imageIndex == 0) MessageBox.Show("這是第一張！");
        }
        private void nextImageShow()
        {
            if (imagesList.Count() == 0) return;
            if (imageIndex > imagesList.Count - 1) { imageIndex = imagesList.Count(); return; }
            while (imageIndex + 1 < imagesList.Count() && !File.Exists(imagesList[++imageIndex])) { }
            //next image
            pictureBox1.ImageLocation = imagesList[imageIndex];
            //顯示圖檔全檔名
            textBox1.Text = pictureBox1.ImageLocation;
            if (imageIndex == imagesList.Count() - 1) MessageBox.Show("這是最後一張！");
        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {//會與Click事件干擾，取消操作，且開啟檔案已設置專屬按鈕，不必於此再操作。有需求再說
            //MouseEventArgs ee = (MouseEventArgs)e;
            //if (pictureBox1.ImageLocation != "")
            //{
            //    switch (ModifierKeys)
            //    {
            //        case Keys.None:
            //            //以Windows系統預設的程式 （「相片檢視器」）開啟圖檔
            //            if (ee.Button == MouseButtons.Left)
            //            {
            //                Process.Start(pictureBox1.ImageLocation);
            //            }
            //            break;
            //        case Keys.Control:
            //            openByMicrosoftOfficePictureManager();
            //            break;

            //        default:
            //            break;
            //    }

            //}

        }

        bool loadImageList(string path)
        {
            if (!Directory.Exists(path))
            {
                MessageBox.Show("資料夾路徑有誤，請重新貼到textBox1(即「要處理的資料夾路徑」文字框)");
                return false;
            }
            sourcePath = path;
            string[] filesList;
            filesList = io.Directory.GetFiles(path);
            if (imagesList.Count() > 0) clearImageList();
            foreach (var fileName in filesList)
            {
                //if ( "jpg,png,bmp,gif,tif".IndexOf(item.Substring(item.Length - 3), 0, 
                //    //不分大小寫
                //    StringComparison.CurrentCultureIgnoreCase) > -1)
                //202301100950 creedit chatGPT ：C# Find Array Element IgnoreCase：
                if (imgExt.Any(x => x.Equals(
                    //chatGPT：C# 取得檔案副檔名：C# 中有一個名為 Path 類別，它提供了用來操作文件路徑和文件名的方法。您可以使用 Path.GetExtension 方法來取得文件的副檔名。
                    Path.GetExtension(fileName), StringComparison.OrdinalIgnoreCase)))
                    imagesList.Add(fileName);

            }
            if (imagesList.Count == 0)
            {
                MessageBox.Show("此資料夾內沒有圖檔！");
                return false;
            }
            else
            {
                //顯示目前正在處理的資料夾路徑
                this.Text = sourcePath;
                //顯示正在顯示處理之檔名（不含路徑）
                label1Text_FileName();
                if (imageIndex != 0) imageIndex = 0;
                return true;
            }
        }
        string newFullName;
        int currentImageIndex;
        void movefileDestination(string sourceFullName, string dirDestination)
        {
            if (dirDestination == sourcePath)
            {
                MessageBox.Show("來源與目的地的路徑相同，請重新指定！", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (File.Exists(sourceFullName) && Directory.Exists(dirDestination))
            {
                string fileName = Path.GetFileName(sourceFullName);//sourceFullName.Substring(sourceFullName.LastIndexOf("\\") + 1);
                dirDestination = dirDestination.Substring(dirDestination.Length - 1, 1) == "\\" ? dirDestination : dirDestination + "\\";
                newFullName = dirDestination + fileName;
                //如果目的地已有同檔名者：
                int i = 0;
                //避免檔名衝突
                while (File.Exists(newFullName))
                {
                    newFullName = dirDestination + Path.GetFileNameWithoutExtension(newFullName) + (i++).ToString() + Path.GetExtension(newFullName);
                }
                if (imageIndex == imagesList.Count() - 1)//取得目前圖檔在清單中的位置,以便/重新載入圖檔清單時顯示被刪除的下一張圖檔，若沒有下一張則顯示前一張
                    currentImageIndex = --imageIndex;
                else
                    currentImageIndex = imageIndex < 0 ? 0 : imageIndex;
                try
                {
                    File.Move(sourceFullName, newFullName);
                }
                catch (Exception)
                {

                    throw;
                }
                //重新載入圖檔清單並顯示被刪除的下一張圖檔，若沒有下一張則顯示前一張
                if (!loadImageList(sourcePath)) return;
                //若有圖檔則載入被刪除的後一張圖檔
                imageIndex = currentImageIndex;
                pictureBox1.ImageLocation = imagesList[currentImageIndex];
                textBox1.Text = pictureBox1.ImageLocation;

            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            openByDllhost();
        }
        bool chkFileExist()
        {
            string xp = textBox1.Text;
            return (xp != "" || File.Exists(xp));
        }
        //以相片檢視器開啟//Default apllication
        void openByDllhost()
        {
            if (chkFileExist())
                Process.Start(textBox1.Text);
            //Process.Start(Environment.GetFolderPath(
            //    Environment.SpecialFolder.Windows) +
            //    "\\System32\\dllhost.exe", textBox1.Text) ;
        }

        //以「Microsoft Office Picture Manager」開啟圖檔
        void openByMicrosoftOfficePictureManager()
        {

            if (chkFileExist())
            {
                //Process.Start(@"C:\Program Files(x86)\Microsoft Office\Office12\OIS.exe",
                //    pictureBox1.ImageLocation);
                string exefile = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) +
                                @"\Microsoft Office\Office12\OIS.exe";
                if (!File.Exists(exefile))
                {
                    MessageBox.Show("並未安裝 Microsoft Office Picture Manager！", "", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button1,
                        //獨占式訊息視窗
                        MessageBoxOptions.DefaultDesktopOnly);
                    return;
                }
                Process.Start(exefile, pictureBox1.ImageLocation);
            }
        }
        //以小畫家開啟
        void openByMsPaint()
        {
            if (chkFileExist())
            {
                Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Windows) +
                    "\\system32\\mspaint.exe",
                    pictureBox1.ImageLocation);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openByMicrosoftOfficePictureManager();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openByMsPaint();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.None)
                textBox_Ciick_importPath_Movefile(ref textBox4);
        }

        private void textBox_Ciick_importPath_Movefile(ref wfrm.TextBox textBox)
        {
            //importPath
            string x = Clipboard.GetText();
            if ((textBox.Text == "" && Directory.Exists(x)) || (Directory.Exists(x) && x != textBox.Text))
            {
                textBox.Text = x; Clipboard.Clear();
            }
            //Movefile
            else
            {
                if (Directory.Exists(textBox.Text) && File.Exists(textBox1.Text))
                    movefileDestination(textBox1.Text, textBox.Text);
            }
        }

        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {
            label1Text_FileName();
        }

        private void label1Text_FileName()
        {
            if (imagesList.Count == 0) return;
            string f = imagesList[imageIndex];
            label1.Text = f.Substring(f.LastIndexOf("\\") + 1);
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            string d = Clipboard.GetText();
            if (textBox1.Text == "要處理的資料夾路徑" && Directory.Exists(d))
                textBox1.Text = d;
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.None)
                textBox_Ciick_importPath_Movefile(ref textBox3);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_MouseDown(object sender, MouseEventArgs e)
        {
            //開啟資料夾以檢視
            if (e.Button == MouseButtons.Left)
                if (ModifierKeys == Keys.Control)
                    selectFileInExplorer(textBox4.Text, newFullName);
        }

        void selectFileInExplorer(string dir, string fullname="")
        {
            if (File.Exists(fullname)) 
            //檔案總管中將已其選取
            //https://www.ruyut.com/2022/05/csharp-open-in-file-explorer.html
            ////把之前開過的關閉
            //if (prcssDownloadImgFullName != null && prcssDownloadImgFullName.HasExited)
            //{
            //    prcssDownloadImgFullName.WaitForExit();
            //    prcssDownloadImgFullName.Close();
            //    ////prcssDownloadImgFullName.WaitForExit();
            //    //prcssDownloadImgFullName.Kill();
            //}
            prcssDownloadImgFullName = System.Diagnostics.Process.Start("Explorer.exe", $"/e, /select ,{fullname}");
            else//沒有指定檔案則開啟資料夾
            {
                if (Directory.Exists(dir))
                    Process.Start("Explorer.exe",dir);
            }
        }

        private void textBox3_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                if (ModifierKeys == Keys.Control)
                    selectFileInExplorer(textBox3.Text, newFullName);
        }
    }
}
