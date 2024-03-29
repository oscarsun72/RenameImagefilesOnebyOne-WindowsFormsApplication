﻿using System;
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
using System.Data.OleDb;
using System.Threading;
using System.Threading.Tasks;

namespace 圖檔重新命名_WindowsFormsApplication1
{
    public partial class Form1 : wfrm.Form
    {
        Process prcssImgFullName;
        //副檔名清單
        readonly string[] imgExt = { ".jpg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".jpeg" };
        string sourcePath = "";
        List<string> imagesList = new List<string>();
        static int xlsRecordCount = 0;
        public Form1()
        {
            InitializeComponent();
            //202301110121 creedit chatGPT：Read Excel into C# code：
            string filePath = "H:\\共用雲端硬碟\\黃老師遠端工作\\沛榮相片整理重新命名用資料表.xlsx";
            //華岡學習雲
            if (!File.Exists(filePath))
                //filePath = "G:\\共用雲端硬碟\\黃老師遠端工作\\沛榮相片整理重新命名用資料表.xlsx";
                filePath = "G:\\共用雲端硬碟\\黃老師遠端工作\\沛榮相片整理重新命名用資料表.xls";
            //黃老師電腦
            if (!File.Exists(filePath))
                filePath = "K:\\共用雲端硬碟\\黃老師遠端工作\\沛榮相片整理重新命名用資料表.xlsx";
            string connectionString;
            if (Path.GetExtension(filePath) != ".xls")
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";" +
                                        "Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";
            else
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";" +
                                            "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT COUNT(*) FROM [工作表1$]";

                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    xlsRecordCount = (int)command.ExecuteScalar();
                }

                query = "SELECT 本家, 家人, 師友生_臺大, 師友生_臺灣學術界, 師友生_其他學術界, 會議, 其他 FROM [工作表1$]";
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        SimpleList family = new SimpleList(), families = new SimpleList(), ntu = new SimpleList()
                        , tw = new SimpleList(), cademia = new SimpleList(), conference = new SimpleList()
                        , others = new SimpleList();
                        // Populate the List
                        family.Add("1本家"); families.Add("2家人"); ntu.Add("3師友生_臺大"); tw.Add("4師友生_臺灣學術界");
                        cademia.Add("5師友生_其他學術界"); conference.Add("6會議 "); others.Add("7其他");
                        while (reader.Read())
                        {
                            if (reader[0].ToString() != "")
                                family.Add(reader[0].ToString());
                            if (reader[1].ToString() != "")
                                families.Add(reader[1].ToString());
                            if (reader[2].ToString() != "")
                                ntu.Add(reader[2].ToString());
                            if (reader[3].ToString() != "")
                                tw.Add(reader[3].ToString());
                            if (reader[4].ToString() != "")
                                cademia.Add(reader[4].ToString());
                            if (reader[5].ToString() != "")
                                conference.Add(reader[5].ToString());
                            if (reader[6].ToString() != "")
                                others.Add(reader[6].ToString());
                        }
                        comboBox1.DataSource = family;
                        comboBox2.DataSource = families;
                        comboBox3.DataSource = ntu;
                        comboBox4.DataSource = tw;
                        comboBox5.DataSource = cademia;
                        comboBox6.DataSource = conference;
                        comboBox7.DataSource = others;
                        comboBox1.MaxDropDownItems = family.Count < 101 ? family.Count : 100;
                        comboBox2.MaxDropDownItems = families.Count < 101 ? families.Count : 100;
                        comboBox3.MaxDropDownItems = ntu.Count < 101 ? ntu.Count : 100;
                        comboBox4.MaxDropDownItems = tw.Count < 101 ? tw.Count : 100;
                        comboBox5.MaxDropDownItems = cademia.Count < 101 ? cademia.Count : 100;
                        comboBox6.MaxDropDownItems = conference.Count < 101 ? conference.Count : 100;
                        comboBox7.MaxDropDownItems = others.Count < 101 ? others.Count : 100;
                        textBox2.Text = "重新命名預覽";
                    }
                }
                //20230114 chatGPT大菩薩：Connection Close In Using：不用再額外手動close，因為在using block結束後會自動close。
                //connection.Close();
            }



            pictureBox1.MouseWheel += new MouseEventHandler(pictureBox1_MouseWheel);

        }

        string renameFilename()
        {//202301110401 creedit chatGPT：重命名檔案名稱：
            string oldfullName = pictureBox1.ImageLocation;
            if (!File.Exists(oldfullName)) return "";
            string newFullName = Path.Combine(Path.GetDirectoryName(oldfullName), textBox2.Text + Path.GetExtension(oldfullName));
            if (newFullName == oldfullName) return "";
            int i = 0;
            while (File.Exists(newFullName))
            {
                newFullName = Path.Combine(Path.GetDirectoryName(oldfullName), textBox2.Text + i++ + Path.GetExtension(oldfullName));
            }
            try
            {
                File.Copy(oldfullName, newFullName);
                File.Delete(oldfullName);
                nextImageShow();
                //開啟檔案總管檢視，重新命名的檔案會被選取
                //System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{newFullName}\"");
            }
            //catch (Exception ex)
            //{
            //    // handle exception
            //}

            //try
            //{
            //    File.Copy(oldfullName, newFullName);
            //    File.Delete(oldfullName);
            //    if(prcssDownloadImgFullName==null)
            //    prcssDownloadImgFullName = System.Diagnostics.Process.Start("Explorer.exe", $"/e, /select ,{newFullName}");
            //    else
            //        prcssDownloadImgFullName.Start("Explorer.exe", $"/e, /select ,{newFullName}");
            //}
            catch (Exception)
            {
                throw;
            }
            form1_colorInfo();
            return newFullName;
        }
        Process prcssNewImgFullName;

        //執行重新命名
        private void button1_Click(object sender, EventArgs e)
        {
            if (ModifierKeys == Keys.None)
            {
                string f = renameFilename();
                if (loadImageList(sourcePath))
                {
                    for (int i = 0; i < imagesList.Count; i++)
                    {
                        if (imagesList[i] == f)
                        {
                            imageIndex = i; break;
                        }
                    }
                    pictureBox1.ImageLocation = f;
                }

            }
            else//開啟檔案總管檢視，重新命名的檔案會被選取
            {
                //把之前開過的關閉
                if (prcssNewImgFullName != null && prcssNewImgFullName.HasExited)
                {
                    prcssNewImgFullName.WaitForExit();
                    prcssNewImgFullName.Close();
                    ////prcssDownloadImgFullName.WaitForExit();
                    //prcssDownloadImgFullName.Kill();
                }

                prcssNewImgFullName = System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{renameFilename()}\"");
            }


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

        }

        //        

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

            switch (e.KeyCode)
            {
                case Keys.Escape:
                    //在編輯時例外
                    if (textBox2.Focused) return;
                    if (MessageBox.Show("結束應用程式？", "確定結束？", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                        == DialogResult.OK)
                        this.Close();
                    break;
                case Keys.F1:
                    Process.Start("https://github.com/oscarsun72/RenameImagefilesOnebyOne-WindowsFormsApplication#readme");
                    break;

                //瀏覽圖檔

                case Keys.Home:
                    //在編輯時例外
                    if (textBox2.Focused) return;
                    //第一張圖檔
                    if (imagesList.Count() == 0) return;
                    imageIndex = 0;
                    pictureBox1.ImageLocation = imagesList[imageIndex];
                    //顯示圖檔全檔名
                    textBox1.Text = pictureBox1.ImageLocation;
                    break;
                case Keys.End:
                    //在編輯時例外
                    if (textBox2.Focused) return;
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
                    //在編輯時例外
                    if (textBox2.Focused) return;
                    nextImageShow();
                    break;
                case Keys.PageDown:
                    nextImageShow();
                    break;
                case Keys.Up:
                    previousImageShow();
                    break;
                case Keys.Left:
                    //在編輯時例外
                    if (textBox2.Focused) return;
                    previousImageShow();
                    break;
                case Keys.PageUp:
                    previousImageShow();
                    break;
                default:
                    break;
            }
        }



        //202301110321 creedit chatGPT：簡化comboBox事件處理：
        private void comboBox_Click_SelectedIndexChanged(object sender, EventArgs e)//(wfrm.ComboBox cmbx)
        {
            string x = ((wfrm.ComboBox)sender).SelectedValue.ToString();
            if (textBox2.Text == "重新命名預覽")
            {
                textBox2.Text = "";
            }
            textBox2.Text += x;

        }

        private void comboBox_Click(object sender, EventArgs e)//(wfrm.ComboBox cmbx)
        {
            ((wfrm.ComboBox)sender).DroppedDown = true;
        }

        class SimpleList : IList //https://docs.microsoft.com/zh-tw/dotnet/api/system.collections.ilist?view=netframework-4.8
        {
            private object[] _contents = new object[xlsRecordCount];//new object[8];
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
                form1_colorInfo();
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
            textBox2Text_FileName();
            if (textBox1.Text != pictureBox1.ImageLocation) textBox1.Text = pictureBox1.ImageLocation;
        }

        private void textBox2Text_FileName()
        {
            if (imagesList.Count == 0) return;
            string f = imagesList[imageIndex];
            textBox2.Text = Path.GetFileNameWithoutExtension(f);
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

        void selectFileInExplorer(string dir, string fullname = "")
        {
            if (File.Exists(fullname))
            {
                //檔案總管中將已其選取
                //https://www.ruyut.com/2022/05/csharp-open-in-file-explorer.html
                //把之前開過的關閉
                if (prcssImgFullName != null && prcssImgFullName.HasExited)
                {
                    prcssImgFullName.WaitForExit();
                    prcssImgFullName.Close();
                    ////prcssDownloadImgFullName.WaitForExit();
                    //prcssDownloadImgFullName.Kill();
                }
                prcssImgFullName = System.Diagnostics.Process.Start("Explorer.exe", $"/e, /select ,{fullname}");
            }
            else//沒有指定檔案則開啟資料夾
            {
                if (Directory.Exists(dir))
                    Process.Start("Explorer.exe", dir);
            }
        }

        private void textBox3_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                if (ModifierKeys == Keys.Control)
                    selectFileInExplorer(textBox3.Text, newFullName);
        }


        /* 實作： 若在「重新命名預覽」想要保留原檔名前綴數字的部分，以便歸類排序尋找篩選……，則可點擊此「方便重新命名檔案」二下，亦會以後綴方式自動填入下方
         * 「重新命名預覽」方塊框中 20230112*/
        private void label1_DoubleClick(object sender, EventArgs e)
        {
            string result = getPrefixNumberNewMethod(label1.Text);
            if (result != "")
            {
                switch (ModifierKeys)
                {   //若按下 Ctrl 再點二下，則會以前綴方式自動填入下方「重新命名預覽」方塊框中
                    case Keys.Control:
                        textBox2.Text = result += "_" + textBox2.Text;
                        break;
                    case Keys.None:
                        Clipboard.SetText(result);
                        break;
                }
            }
        }

        private string getPrefixNumberNewMethod(string input)
        {
            string result = "";
            //creedit chatGPT：
            //bool isNum = true;
            for (int i = 0; i < input.Length; i++)
            {
                if (!Char.IsDigit(input[i]) && input[i] != '.')
                {
                    //isNum = false;
                    break;
                }
                result += input[i];
            }

            return result;
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Left:
                    switch (ModifierKeys)
                    {
                        //實作： 若按下前按住 Ctrl 鍵，則可直接貼入下方的「重新命名預覽」作為原有文字之後綴，以便編輯。
                        case Keys.Control:
                            if (label1.Text != "方便重新命名檔案")
                            {
                                if (!isDoubleClick(DateTime.Now))
                                {
                                    string fn = Path.GetFileNameWithoutExtension(label1.Text);
                                    if (textBox2.Text.IndexOf(fn) == -1)
                                        textBox2.Text += fn;
                                }
                                else
                                {
                                    //label1_DoubleClick(sender, e);
                                    return;
                                }
                            }
                            break;
                        //實作： 在「方便重新命名檔案」處若是顯示為正在檢視的圖片檔名（含副檔名），則滑鼠左鍵點一下，即可將其檔名（不含副檔名）的部分複製到剪貼簿以備用。
                        case Keys.None:
                            if (label1.Text != "方便重新命名檔案")
                                Clipboard.SetText(Path.GetFileNameWithoutExtension(label1.Text));
                            break;
                    }
                    break;
                case MouseButtons.None:
                    break;
                case MouseButtons.Right:
                    break;
                case MouseButtons.Middle:
                    break;
                case MouseButtons.XButton1:
                    break;
                case MouseButtons.XButton2:
                    break;
                default:
                    break;
            }

        }


        //判斷是否是 DoubleClick事件,creedit chatGPT：Determine DoubleClick Event Called：
        private DateTime lastClick = DateTime.Now;
        bool isDoubleClick(DateTime now)
        {
            TimeSpan clickInterval = now - lastClick;
            if (clickInterval.TotalMilliseconds < SystemInformation.DoubleClickTime)//使用 SystemInformation.DoubleClickTime 屬性來判斷是否為 DoubleClick 事件。
            {
                // it's a double click
                return true;
            }
            else
            {
                // it's a single click
                lastClick = DateTime.Now;
                return false;
            }
        }

        //textBox2=重新命名預覽
        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            //textBox2.SelectedText = "";
        }

        wfrm.ToolTip toolTip1 = new wfrm.ToolTip();
        private void label1_MouseMove(object sender, MouseEventArgs e)
        {// creedit with YouChat：20230112
            //toolTip1.SetToolTip(label1, "3q");
            tooltipConstructor(sender,
             @"滑鼠左鍵點2下，可將此檔名（不含副檔名）的部分複製到剪貼簿以備用。
                          >>  若按下前按住 Ctrl 鍵，則可直接貼入下方的「重新命名預覽」作為原有文字之後綴，以便編輯。
                    >> 按二下，會複製原檔名前綴數字的部分到剪貼簿，
                           >>若按下 Ctrl 再點二下，則會以前綴+底線「_」方式自動填入下方「重新命名預覽」方塊框中。
                    （通常由數位器材攝製的，是以序號方式自動命名，是亦存有影像原始攝製之時間及群組的特性，往往有助於辨識年久不認的往事及當時攝製之場景。）");

        }

        private void textBox2_MouseMove(object sender, MouseEventArgs e)
        {
            string tooltipText =
                @"若按下 Ctrl 再點一下，則會只留下原檔名前綴的數字+底線「_」部分以備用
                    （通常由數位器材攝製的，是以序號方式自動命名，是亦存有影像原始攝製之時間及群組的特性，往往有助於辨識年久不認的往事及當時攝製之場景）
                  若空白時在此點一下滑鼠，則會輸入剪貼簿裡的文字";
            tooltipConstructor(sender, tooltipText);

        }

        private void tooltipConstructor(object sender, string tooltipText)
        {
            if (toolTip1.GetToolTip((Control)sender) != tooltipText)
                toolTip1.SetToolTip((Control)sender, tooltipText);
        }

        //重新命名預覽
        private void textBox2_MouseDown(object sender, MouseEventArgs e)
        {
            switch (ModifierKeys)
            {
                //按下Ctrl 再點滑鼠
                case Keys.Control:
                    string x = getPrefixNumberNewMethod(label1.Text);
                    if (x != "")
                        textBox2.Text = x + "_";
                    break;
                case Keys.None:
                    //點滑鼠一下
                    if (textBox2.Text == "")
                        textBox2.Text = Clipboard.GetText();
                    break;
                default:
                    break;
            }
        }

        void textBox3_tooltip(object sender)
        {
            tooltipConstructor(sender,
                @"● 初次點擊「要移動到的資料夾路徑」，若剪貼簿裡已複製有資料夾路徑，則會直接輸入。
                ● 當「要移動到的資料夾路徑」文字框內是有效資料夾路徑時，點擊一次，即會將現前的圖檔移至框中指定的目錄去，並重新載入圖檔清單，顯示被移動的下一個圖檔。
                    一般可以在重新命名後想改放位置，再做此指定操作。
                ● 若「要移動到的資料夾路徑」框中已是資料夾路徑，而不想行動，請對框按滑鼠右鍵（千萬別按左鍵，否則即成移動），
                    然後再以按鍵輸入，或選擇跳出的快顯功能表菜單中「刪除」以清除路徑。
                ● 總之，「要移動到的資料夾路徑」諸框中，只要是路徑，一經左鍵點擊，即會將現前圖檔移動到該路徑下。千萬注意！！
                ● 若想要檢閱所移動的圖檔，可於按下左鍵點擊「要移動到的資料夾路徑」框前，按住 Ctrl ，則會在移動完畢後開啟一個新的檔案總管並選定此檔以供檢視。（Ctrl 就是掌握的意思，不想被動，想主控一切，就記得按住 control ！ ^_^ ）");

        }
        private void textBox3_MouseMove(object sender, MouseEventArgs e)
        {
            textBox3_tooltip(sender);
        }

        private void textBox4_MouseMove(object sender, MouseEventArgs e)
        {
            textBox3_tooltip(sender);
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            string x = "";
            switch (ModifierKeys)
            {
                //若剪貼簿裡的內容是有效路徑且此框內所顯示者不同，則以滑鼠點一下，會自動貼入剪貼簿裡的路徑
                case Keys.None:
                    x = Clipboard.GetText();
                    if (Directory.Exists(x) && x != textBox1.Text)
                        textBox1.Text = Clipboard.GetText();
                    break;
                //若按住 Ctrl 再點一下，則會開啟此框所顯示的路徑資料夾以供檢視（若是全檔名則開啟所在資料夾）
                case Keys.Control:
                    x = textBox1.Text;
                    if (File.Exists(x)) Process.Start(Path.GetDirectoryName(x));
                    else if (Directory.Exists(x)) Process.Start(x);
                    break;

            }
        }

        void form1_colorInfo()
        {//凡更動過後都當如此
         //- 移動過後則表單會變色一下（閃一下）以示- 命名過後則表單會變色一下（閃一下）以示
            Color c = BackColor;
            BackColor = Color.Tan;
            Refresh();
            //Thread.Sleep(1700);//這才有效（這是管當前執行緒）
            //Task.Delay(13950);//這無效，管的不是這裡，當是另一個執行緒
            BackColor = c;
            Refresh();
        }

        private void textBox1_MouseMove(object sender, MouseEventArgs e)
        {
            tooltipConstructor(sender,
                @"「要處理的資料夾路徑」文字框：作為首先讀取要操作的資料夾路徑輸入用，輸入後會擷取資料夾內所有圖檔（不含子資料夾），
                        並會即刻顯示第一張圖檔以供檢視操作，且在此方塊框中改以顯示此當前圖檔的全檔名（目錄路徑+檔名+副檔名），可供後事複製等備用。
                - 若剪貼簿裡的內容是有效路徑且此框內所顯示者不同，則以滑鼠點一下，會自動貼入剪貼簿裡的路徑；
                - 若按住 Ctrl 再點一下，則會開啟此框所顯示的路徑資料夾以供檢視（若是全檔名則開啟所在資料夾）");

        }
    }
}
