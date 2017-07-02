using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Media;

namespace Cari
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region değişkenler
        OleDbConnection conn;
        OleDbConnection cnn;
        string dosya_yolu;
        string hedef,kaynak;
        bool web,promosyon,reklam;
        ListBox dosya_yolları = new ListBox();

        TextBox klasor = new TextBox();
        TextBox id = new TextBox();
        TextBox saat = new TextBox();

        TextBox saat_1 = new TextBox();
        TextBox dakika = new TextBox();

        TextBox saat_n = new TextBox();
        TextBox dakika_n = new TextBox();
        
        ComboBox isler = new ComboBox();
        ComboBox idler = new ComboBox();

        int sa_n, dk_n, sa, dk;
        int sa_fark, dk_fark;

        int y_ekseni = -20;

        #endregion

        #region Fonksiyonlar

        private void is_getir()
        {
            try
            {
                conn.Open();
                OleDbCommand com = new OleDbCommand("select hizmet from hizmet ", conn);
                OleDbDataReader dr = com.ExecuteReader();
                while (dr.Read())
                {
                    hizmet.Items.Add(dr["hizmet"]);

                }
            }
            catch 
            {
                MessageBox.Show("Bir Hata Oluştu","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

        }

        private void boyutlandır()
        {


            try
            {
                DG1.Height = tabControl1.Height - (groupBox1.Height + groupBox2.Height);
                pictureBox1.Width = (tabControl1.Width / 2) + 100;
                pictureBox1.Dock = DockStyle.Right;


                if (treeView1.Width > 100)
                {
                    treeView1.Width = (tabControl1.Width / 2) - 100;
                    treeView1.Dock = DockStyle.Left;
                }
                else
                {
                    treeView1.Width = (tabControl1.Width / 2) + 100;
                    treeView1.Dock = DockStyle.Left;
                }

                richTextBox1.Width = (tabControl1.Width / 2) + 100;
                richTextBox1.Dock = DockStyle.Right;

                dataGridView2.Width = (tabControl1.Width / 2) + 100;
                dataGridView2.Dock = DockStyle.Right;

                webBrowser1.Width = (tabControl1.Width / 2) + 100;
                webBrowser1.Dock = DockStyle.Right;

            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dosya_sec(ListBox list1)
        {
            try
            {
                openFileDialog1.ShowDialog();

                foreach (String file in openFileDialog1.FileNames)
                {

                    list1.Items.Add(file);

                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dosya_ekle(TextBox firma_adi, TextBox dosya_yolu_text, ListBox dosyalar)
        {
            try
            {
                string uzantı1, d_adi, tmp;
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select f_dosya_yolu from firma where f_adi='" + firma_adi.Text + "'", conn);
                adtr1.Fill(dtst1, "firma");


                dosya_yolu_text.Text = "";
                dosya_yolu_text.DataBindings.Add("text", dtst1, "firma.f_dosya_yolu");
                tmp = dosya_yolu_text.Text;
                dosya_yolu_text.DataBindings.Clear();



                for (int i = 0; i < dosyalar.Items.Count; i++)
                {
                    hedef = dosya_yolu_text.Text;

                    dosyalar.SelectedIndex = i;
                    kaynak = dosyalar.SelectedItem.ToString();

                    d_adi = Path.GetFileName(kaynak);
                    uzantı1 = Path.GetExtension(kaynak);


                    if (uzantı1 == ".jpg" || uzantı1 == ".jpeg" || uzantı1 == ".png" || uzantı1 == ".bmp")
                    {
                        dosya_yolu_text.Clear();
                        dosya_yolu_text.Text = tmp + "\\Resim\\";
                        hedef = dosya_yolu_text.Text;
                    }

                    else
                    {
                        dosya_yolu_text.Clear();
                        dosya_yolu_text.Text = tmp + "\\Döküman\\";
                        hedef = dosya_yolu_text.Text;
                    }

                    FileInfo DosyaAdi = new FileInfo(kaynak);
                    if (DosyaAdi.Exists)
                    {
                        File.Copy(kaynak, hedef + d_adi, true);

                        //MessageBox.Show("Dosya Kopyalama İşlemi Gerçekleştirildi");
                    }
                    else
                    {
                        MessageBox.Show("Dosya Kopyalanamadı", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }

                MessageBox.Show("Dosya Kopyalama İşlemi Gerçekleştirildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dosya_sil(TextBox firma_adi, TextBox dosya_yolu_text, ListBox dosyalar)
        {
            try
            {
                string uzantı1, d_adi, tmp;
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select f_dosya_yolu from firma where f_adi='" + firma_adi.Text + "'", conn);
                adtr1.Fill(dtst1, "firma");


                dosya_yolu_text.Text = "";
                dosya_yolu_text.DataBindings.Add("text", dtst1, "firma.f_dosya_yolu");
                tmp = dosya_yolu_text.Text;
                dosya_yolu_text.DataBindings.Clear();



                for (int i = 0; i < dosyalar.Items.Count; i++)
                {
                    hedef = dosya_yolu_text.Text;

                    dosyalar.SelectedIndex = i;
                    kaynak = dosyalar.SelectedItem.ToString();

                    d_adi = Path.GetFileName(kaynak);
                    uzantı1 = Path.GetExtension(kaynak);


                    if (uzantı1 == ".jpg" || uzantı1 == ".jpeg" || uzantı1 == ".png" || uzantı1 == ".bmp")
                    {
                        dosya_yolu_text.Clear();
                        dosya_yolu_text.Text = tmp + "\\Resim\\";
                        hedef = dosya_yolu_text.Text;
                    }

                    else
                    {
                        dosya_yolu_text.Clear();
                        dosya_yolu_text.Text = tmp + "\\Döküman\\";
                        hedef = dosya_yolu_text.Text;
                    }

                    FileInfo DosyaAdi = new FileInfo(kaynak);
                    if (DosyaAdi.Exists)
                    {
                        File.Delete(hedef + d_adi);

                        //MessageBox.Show("Dosya Kopyalama İşlemi Gerçekleştirildi");
                    }
                    else
                    {
                        MessageBox.Show("Dosyalar Silinemedi", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }

                MessageBox.Show("Dosya Silme İşlemi Gerçekleştirildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void firma_Getir(DataGridView dg)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma ", conn);
                adtr1.Fill(dtst1, "firma");
                dg.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void temizle()
        {
            try
            {
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                textBox8.Clear();
                hizmet.Text = "İş Türü Seçiniz";
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void firma_ara()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_adi like'%" + textBox9.Text + "%'", conn);
                adtr1.Fill(dtst1, "firma");
                DG1.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void firma_show(TextBox firma1, TextBox firma2, TextBox firma_adi, TextBox firma_sor, TextBox firma_tel, TextBox firma_fax, TextBox firma_mail, TextBox firma_adres, TextBox firma_vd, TextBox firma_vn, TextBox firma_adi2, ComboBox firma_hizmet, TextBox firma_dosya)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();

                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (firma1.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From firma  where f_adi='" + firma2.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From firma  where f_adi='" + firma1.Text + "' ", conn);
                }
                adtr2.Fill(dtst2, "firma");

                firma_adi.Clear();
                firma_adi.DataBindings.Clear();
                firma_adi.DataBindings.Add("text", dtst2, "firma.f_adi");


                firma_sor.Clear();
                firma_sor.DataBindings.Clear();
                firma_sor.DataBindings.Add("text", dtst2, "firma.f_sorumlusu");


                firma_tel.Clear();
                firma_tel.DataBindings.Clear();
                firma_tel.DataBindings.Add("text", dtst2, "firma.f_tel");

                firma_fax.Clear();
                firma_fax.DataBindings.Clear();
                firma_fax.DataBindings.Add("text", dtst2, "firma.f_fax");

                firma_mail.Clear();
                firma_mail.DataBindings.Clear();
                firma_mail.DataBindings.Add("text", dtst2, "firma.f_mail");

                firma_adres.Clear();
                firma_adres.DataBindings.Clear();
                firma_adres.DataBindings.Add("text", dtst2, "firma.f_adres");


                firma_vd.Clear();
                firma_vd.DataBindings.Clear();
                firma_vd.DataBindings.Add("text", dtst2, "firma.f_vergi_dairesi");


                firma_vn.Clear();
                firma_vn.DataBindings.Clear();
                firma_vn.DataBindings.Add("text", dtst2, "firma.f_vergi_no");


                firma_hizmet.Text = "Hizmet";
                firma_hizmet.DataBindings.Clear();
                firma_hizmet.DataBindings.Add("text", dtst2, "firma.f_hizmet");

                firma_dosya.Clear();
                firma_dosya.DataBindings.Clear();
                firma_dosya.DataBindings.Add("text", dtst2, "firma.f_dosya_yolu");

                firma_adi2.Text = firma_adi.Text;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void hizmet_temizle(Control ctl, Control ctl1, Control ctl2, Control ctl3, Control ctl4)
        {
            try
            {
                foreach (Control c in ctl.Controls)
                {
                    if (c is TextBox)
                    {
                        ((TextBox)c).Clear();
                    }
                    else if (c is ComboBox)
                    {
                        c.Text = "HİZMET";
                    }

                }
                foreach (Control c1 in ctl1.Controls)
                {
                    if (c1 is TextBox)
                    {
                        ((TextBox)c1).Clear();
                    }
                    else if (c1 is ComboBox)
                    {
                        c1.Text = "HİZMET";
                    }

                }
                foreach (Control c2 in ctl2.Controls)
                {
                    if (c2 is TextBox)
                    {
                        ((TextBox)c2).Clear();
                    }
                    else if (c2 is ComboBox)
                    {
                        c2.Text = "HİZMET";
                    }

                }
                foreach (ListBox c3 in ctl3.Controls)
                {

                    c3.Items.Clear();
                }
                foreach (ListBox c4 in ctl4.Controls)
                {

                    c4.Items.Clear();
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dosya_getir(ListBox list1, TextBox text)
        {
            try
            {
                list1.Items.Clear();
                string yol = System.IO.Path.Combine();

                yol = text.Text + "\\Döküman\\";
                foreach (string dosya in System.IO.Directory.GetFiles(yol))
                {
                    list1.Items.Add((dosya));

                }

                yol = text.Text + "\\Resim\\";

                foreach (string dosya1 in System.IO.Directory.GetFiles(yol))
                {
                    list1.Items.Add((dosya1));
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void is_tablosu_getir(DataGridView dg, ComboBox cb)
        {
            try
            {
                conn.Close();
                conn.Open();
                OleDbCommand com = new OleDbCommand("select Person_Adi from personel ", conn);
                OleDbDataReader dr = com.ExecuteReader();
                cb.Items.Clear();
                while (dr.Read())
                {
                    cb.Items.Add(dr["Person_Adi"]);

                }

                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From isler where i_durum='0'  ", conn);
                adtr1.Fill(dtst1, "isler");
                dg.DataSource = dtst1.Tables["isler"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void ise_gonder(DataGridView dg)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From isler where i_durum='0' AND i_tarih='"+maskedTextBox6.Text+"' ", conn);
                adtr1.Fill(dtst1, "isler");
                dg.DataSource = dtst1.Tables["isler"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }
        private void secimi_iptal(Control ctl)
        {
            try
            {
                foreach (Control c in ctl.Controls)
                {
                    if (c is TextBox)
                    {
                        ((TextBox)c).Clear();
                    }
                    else if (c is MaskedTextBox)
                    {
                        ((MaskedTextBox)c).Clear();
                    }
                    else if (c is ComboBox)
                    {
                        c.Text = "PERSONEL SEÇİNİZ";
                    }

                }
                textBox62.Text = "";
                textBox60.Enabled = true;
                button29.Enabled = true;
                comboBox5.Text = "00";
                comboBox6.Text = "00";
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void web_getir()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_hizmet='WEB' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG2.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
            }
            
            catch 
            {
                MessageBox.Show("Bir Hata Oluştu","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }
        private void promosyon_getir()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_hizmet='PROMOSYON' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG3.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void reklam_getir()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_hizmet='REKLAM' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG4.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void web_kont()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();
                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox21.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From web  where w_f_adi='" + textBox13.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From web  where w_f_adi='" + textBox21.Text + "' ", conn);
                }

                adtr2.Fill(dtst2, "web");
                textBox30.Clear();
                textBox30.DataBindings.Clear();
                textBox30.DataBindings.Add("text", dtst2, "web.w_id");

                if (textBox30.Text != "")
                {
                    web = true;
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        private void promosyon_kont()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();
                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox21.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From promosyon  where p_f_adi='" + textBox42.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From promosyon  where p_f_adi='" + textBox43.Text + "' ", conn);
                }

                adtr2.Fill(dtst2, "promosyon");
                textBox33.Clear();
                textBox33.DataBindings.Clear();
                textBox33.DataBindings.Add("text", dtst2, "promosyon.p_id");

                if (textBox33.Text != "")
                {
                    promosyon = true;
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        private void reklam_kont()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();
                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox21.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From reklam  where r_f_adi='" + textBox56.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From reklam  where r_f_adi='" + textBox57.Text + "' ", conn);
                }

                adtr2.Fill(dtst2, "reklam");
                textBox46.Clear();
                textBox46.DataBindings.Clear();
                textBox46.DataBindings.Add("text", dtst2, "reklam.r_id");

                if (textBox46.Text != "")
                {
                    reklam = true;
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void web_show()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();

                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox21.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From web  where w_f_adi='" + textBox13.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From web  where w_f_adi='" + textBox21.Text + "' ", conn);
                }
                adtr2.Fill(dtst2, "web");

                textBox14.Clear();
                textBox14.DataBindings.Clear();
                textBox14.DataBindings.Add("text", dtst2, "web.w_panel");


                textBox15.Clear();
                textBox15.DataBindings.Clear();
                textBox15.DataBindings.Add("text", dtst2, "web.w_mail");


                textBox16.Clear();
                textBox16.DataBindings.Clear();
                textBox16.DataBindings.Add("text", dtst2, "web.w_sifre");

                textBox17.Clear();
                textBox17.DataBindings.Clear();
                textBox17.DataBindings.Add("text", dtst2, "web.w_alan_adi");

                textBox18.Clear();
                textBox18.DataBindings.Clear();
                textBox18.DataBindings.Add("text", dtst2, "web.w_hosting");

                textBox19.Clear();
                textBox19.DataBindings.Clear();
                textBox19.DataBindings.Add("text", dtst2, "web.w_aciklama");
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        private void promosyon_show()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();

                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox21.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From promosyon  where p_f_adi='" + textBox50.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From promosyon  where p_f_adi='" + textBox43.Text + "' ", conn);
                }
                adtr2.Fill(dtst2, "promosyon");

                textBox44.Clear();
                textBox44.DataBindings.Clear();
                textBox44.DataBindings.Add("text", dtst2, "promosyon.p_urun_bilgi");
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void reklam_show()
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();

                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (textBox57.Text == "")
                {
                    adtr2 = new OleDbDataAdapter("select * From reklam  where r_f_adi='" + textBox59.Text + "' ", conn);
                }
                else
                {
                    adtr2 = new OleDbDataAdapter("select * From reklam  where r_f_adi='" + textBox57.Text + "' ", conn);
                }
                adtr2.Fill(dtst2, "reklam");

                textBox58.Clear();
                textBox58.DataBindings.Clear();
                textBox58.DataBindings.Add("text", dtst2, "reklam.r_bilgi");
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void texte_yaz()
        {
            try
            {
                richTextBox2.SelectionAlignment = HorizontalAlignment.Center;
                richTextBox2.Text = "";
                richTextBox2.Text = ("İŞ BİLGİLERİ" + "\n\n" + "Firma: " + textBox75.Text + "\n" + "Hizmet Türü: " + textBox73.Text + "\n" + "Yapılacak İş:" + textBox74.Text + "\n" + "Tarih: " + maskedTextBox3.Text + "\n" + "Saat: " + maskedTextBox4.Text + "\n\n" + "FİRMA BİLGİLERİ" + "\n" + "Firma :" + textBox84.Text + "\n" + "Sorumlu: " + textBox85.Text + "\n" + "Telefon: " + textBox86.Text + "\n" + "Fax: " + textBox87.Text + "\n" + "Email: " + textBox88.Text + "\n" + "Adres: " + textBox89.Text + "\n\n" + "PERSONEL BİLGİLERİ" + "\n" + "Adı: " + textBox77.Text + "\n" + "Telefon: " + textBox78.Text + "\n" + "Mail: " + textBox79.Text);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void is_temizle()
        {
            try
            {
                textBox76.Clear();
                textBox77.Clear();
                textBox78.Clear();
                textBox79.Clear();
                textBox64.Clear();
                textBox75.Clear();
                textBox74.Clear();
                textBox73.Clear();
                maskedTextBox3.Clear();
                maskedTextBox4.Clear();
                textBox82.Clear();
                textBox80.Clear();
                textBox75.Clear();
                textBox84.Clear();
                textBox85.Clear();
                textBox86.Clear();
                textBox87.Clear();
                textBox88.Clear();
                textBox89.Clear();
                textBox90.Clear();
                textBox91.Clear();
                textBox75.Clear();
                comboBox8.Text = "";
                comboBox10.Text = "";
                textBox81.Clear();
                richTextBox2.Clear();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion   

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                tabControl1.Visible = false;
                panel4.Visible = true;

                conn = new OleDbConnection("Provider=Microsoft.Ace.oledb.12.0;data source= cari.mdb");


                boyutlandır();
                is_getir();
                firma_Getir(DG1);
                temizle();

                panel4.Left = panel4.Left + (this.Width / 2) - (panel4.Width / 2);
                panel4.Top = panel4.Top + (this.Height / 2) - (panel4.Height);

                panel3.Left = panel3.Left + (this.Width / 2) - (panel3.Width / 2);
                panel3.Top = panel3.Top + (tabControl1.Height / 2) - (panel3.Height);

                groupBox1.Left = groupBox1.Left + (tabControl1.Width / 2) - (groupBox1.Width / 2);

                groupBox2.Left = groupBox2.Left + (tabControl1.Width / 2) - (groupBox2.Width / 2);
                groupBox3.Left = groupBox3.Left + (tabControl1.Width / 2) - (groupBox3.Width / 2);
                maskedTextBox6.Text = DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
         }
     
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                e.Node.Nodes.Clear();

                FileAttributes attr = File.GetAttributes(e.Node.Text);

                //Dizin mi, dosya mı bakılıyor
                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    foreach (var drive in Directory.GetDirectories(e.Node.Text))
                    {

                        e.Node.Nodes.Add(drive);


                    }

                    foreach (var drive in Directory.GetFiles(e.Node.Text))
                    {

                        e.Node.Nodes.Add(drive);


                    }
                    dosya_yolu = e.Node.Text;
                    dosya_yolu = Path.GetFullPath(dosya_yolu);

                    if (richTextBox1.Visible == true)
                    {
                        richTextBox1.Visible = false;
                    }
                    if (pictureBox1.Visible == true)
                    {
                        pictureBox1.Visible = false;
                    }
                    if (dataGridView2.Visible == true)
                    {
                        dataGridView2.Visible = false;
                    }
                    if (webBrowser1.Visible == false)
                    {
                        webBrowser1.Visible = true;
                    }

                    webBrowser1.Url = new Uri(dosya_yolu);
                }
                else
                {
                    dosya_yolu = e.Node.Text;
                    textBox12.Text = dosya_yolu;
                    dosya_yolu = Path.GetFullPath(dosya_yolu);

                    string uzantı;

                    uzantı = Path.GetExtension(dosya_yolu);//Belli Bir Karakterden Sonrasını Alma
                    //Karakterden Sonraki Bölüm

                    richTextBox1.Clear();

                    if (uzantı == ".jpg" || uzantı == ".jpeg" || uzantı == ".png" || uzantı == ".bmp")
                    {
                        if (richTextBox1.Visible == true)
                        {
                            richTextBox1.Visible = false;
                        }
                        if (pictureBox1.Visible == false)
                        {
                            pictureBox1.Visible = true;
                        }
                        if (webBrowser1.Visible == true)
                        {
                            webBrowser1.Visible = false;
                        }
                        if (dataGridView2.Visible == true)
                        {
                            dataGridView2.Visible = false;
                        }
                        pictureBox1.Image = Image.FromFile(dosya_yolu);


                    }


                    else if (uzantı == ".doc" || uzantı == ".docx")
                    {

                        Microsoft.Office.Interop.Word.Application wd = new Microsoft.Office.Interop.Word.Application();

                        wd.Documents.Open(dosya_yolu);
                        wd.Selection.WholeStory();
                        wd.Selection.Copy();
                        richTextBox1.Paste();
                        wd.Quit();

                        if (richTextBox1.Visible == false)
                        {
                            richTextBox1.Visible = true;
                        }
                        if (dataGridView2.Visible == true)
                        {
                            dataGridView2.Visible = false;
                        }
                        if (webBrowser1.Visible == true)
                        {
                            webBrowser1.Visible = false;
                        }
                        if (pictureBox1.Visible == true)
                        {
                            pictureBox1.Visible = false;
                        }

                    }
                    else if (uzantı == ".xls" || uzantı == ".xlsx")
                    {

                        if (richTextBox1.Visible == true)
                        {
                            richTextBox1.Visible = false;
                        }
                        if (pictureBox1.Visible == true)
                        {
                            pictureBox1.Visible = false;
                        }
                        if (webBrowser1.Visible == true)
                        {
                            webBrowser1.Visible = false;
                        }
                        if (dataGridView2.Visible == false)
                        {
                            dataGridView2.Visible = true;
                        }

                        // conn = new OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source= cari.mdb");


                        OleDbConnectionStringBuilder excelAyar = new OleDbConnectionStringBuilder();
                        excelAyar.DataSource = dosya_yolu; // excel kitabının tam yol adı
                        excelAyar.Provider = "Microsoft.ace.OLEDB.12.0";
                        excelAyar["Extended Properties"] = "Excel 12.0 Xml;HDR=YES";
                        string excelSayfaAdi = "Sayfa1"; // verileri alacağınız Excel sayfasının adı
                        OleDbConnection excelBag = new OleDbConnection(excelAyar.ConnectionString);

                        excelBag.Open();
                        OleDbDataAdapter adap = new OleDbDataAdapter("SELECT * FROM [" + excelSayfaAdi + "$]", excelBag);
                        DataTable dt = new DataTable(); adap.Fill(dt);
                        dataGridView2.DataSource = dt;


                    }
                    else if (uzantı == ".pdf")
                    {

                        if (richTextBox1.Visible == true)
                        {
                            richTextBox1.Visible = false;
                        }
                        if (pictureBox1.Visible == true)
                        {
                            pictureBox1.Visible = false;
                        }
                        if (dataGridView2.Visible == true)
                        {
                            dataGridView2.Visible = false;
                        }
                        if (webBrowser1.Visible == false)
                        {
                            webBrowser1.Visible = true;
                        }

                        webBrowser1.Url = new Uri(dosya_yolu);

                    }
                    else
                    {
                        if (richTextBox1.Visible == false)
                        {
                            richTextBox1.Visible = true;
                        }
                        if (dataGridView2.Visible == true)
                        {
                            dataGridView2.Visible = false;
                        }
                        if (webBrowser1.Visible == true)
                        {
                            webBrowser1.Visible = false;
                        }
                        if (pictureBox1.Visible == true)
                        {
                            pictureBox1.Visible = false;
                        }


                        richTextBox1.LoadFile(dosya_yolu, RichTextBoxStreamType.PlainText);
                    }

                }

            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
         
            
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            try
            {
                boyutlandır();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DG1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox1.Enabled = false;
                button1.Enabled = false;
                textBox1.Text = (DG1.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox2.Text = (DG1.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox3.Text = (DG1.Rows[e.RowIndex].Cells[3].Value.ToString());
                textBox4.Text = (DG1.Rows[e.RowIndex].Cells[4].Value.ToString());
                textBox5.Text = (DG1.Rows[e.RowIndex].Cells[5].Value.ToString());
                textBox6.Text = (DG1.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox7.Text = (DG1.Rows[e.RowIndex].Cells[7].Value.ToString());
                textBox8.Text = (DG1.Rows[e.RowIndex].Cells[8].Value.ToString());
                hizmet.Text = (DG1.Rows[e.RowIndex].Cells[9].Value.ToString());
                klasor.Text = (DG1.Rows[e.RowIndex].Cells[10].Value.ToString());

                dosya_getir(dosya_yolları, klasor);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand com = new OleDbCommand("insert into firma(f_adi,f_sorumlusu,f_tel,f_fax,f_mail,f_adres,f_vergi_dairesi,f_vergi_no,f_hizmet,f_dosya_yolu) values ('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + textBox4.Text + "','" + textBox5.Text + "','" + textBox6.Text + "','" + textBox7.Text + "','" + textBox8.Text + "','" + hizmet.Text + "','" + "Root\\" + textBox1.Text + "')", conn);
                MessageBox.Show("Kayıt Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                com.Connection.Open();
                com.ExecuteNonQuery();
                conn.Close();

                string folderName = @"Root";
                string pathString = System.IO.Path.Combine(folderName, textBox1.Text);
                string pathString1 = System.IO.Path.Combine(folderName, textBox1.Text, "Resim");
                string pathString2 = System.IO.Path.Combine(folderName, textBox1.Text, "Döküman");
                Directory.CreateDirectory(pathString);
                Directory.CreateDirectory(pathString1);
                Directory.CreateDirectory(pathString2);

                temizle();
                firma_Getir(DG1);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            try
            {
                firma_ara();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                firma_ara();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update firma set f_adi= '" + textBox1.Text + "', f_sorumlusu= '" + textBox2.Text + "',f_tel='" + textBox3.Text + "',f_fax= '" + textBox4.Text + "',f_mail= '" + textBox5.Text + "',f_adres= '" + textBox6.Text + "',f_vergi_dairesi= '" + textBox7.Text + "',f_vergi_no= '" + textBox8.Text + "',f_hizmet= '" + hizmet.Text + "' where f_adi = '" + textBox1.Text + "'", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();

                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                temizle();
                textBox1.Enabled = true;
                firma_Getir(DG1);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Bu İşlem Sadece Eklenen Firmaları Siler \n\nSilmek İstediğinizden Eminmisiniz", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    conn.Close();
                    // seçili olan kayıdı siler
                    OleDbCommand sil = new OleDbCommand("delete from firma where f_adi= '" + textBox1.Text + "' ", conn);
                    sil.Connection.Open();
                    sil.ExecuteNonQuery();

                    conn.Close();
                    // seçili olan kayıdı siler
                    OleDbCommand sil1 = new OleDbCommand("delete from web where w_f_adi= '" + textBox1.Text + "' ", conn);
                    sil1.Connection.Open();
                    sil1.ExecuteNonQuery();

                    conn.Close();
                    // seçili olan kayıdı siler
                    OleDbCommand sil2 = new OleDbCommand("delete from promosyon where p_f_adi= '" + textBox1.Text + "' ", conn);
                    sil2.Connection.Open();
                    sil2.ExecuteNonQuery();

                    conn.Close();
                    // seçili olan kayıdı siler
                    OleDbCommand sil3 = new OleDbCommand("delete from reklam where r_f_adi= '" + textBox1.Text + "' ", conn);
                    sil3.Connection.Open();
                    sil3.ExecuteNonQuery();

                    conn.Close();
                    if (MessageBox.Show("Silme İşlemi Tamamlandı \n\nDosyaları Silmek İster Misiniz?", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {

                        Directory.Delete(klasor.Text, true);

                    }
                    MessageBox.Show("Silme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    temizle();
                    textBox1.Enabled = true;
                    firma_Getir(DG1);
                }
                else
                {
                    MessageBox.Show("Silme İşlemi İptal Edildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                temizle();
                textBox1.Enabled = true;
                button1.Enabled = true;
                firma_Getir(DG1);
                textBox9.Clear();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            try
            {
                firma_Getir(DG5);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma where f_adi='" + textBox10.Text + "'", conn);
                adtr1.Fill(dtst1, "firma");
                textBox11.Clear();
                textBox11.DataBindings.Clear();
                textBox11.DataBindings.Add("text", dtst1, "firma.f_dosya_yolu");

                textBox72.Clear();
                textBox72.DataBindings.Clear();
                textBox72.DataBindings.Add("text", dtst1, "firma.f_adi");

                textBox65.Clear();
                textBox65.DataBindings.Clear();
                textBox65.DataBindings.Add("text", dtst1, "firma.f_sorumlusu");

                textBox66.Clear();
                textBox66.DataBindings.Clear();
                textBox66.DataBindings.Add("text", dtst1, "firma.f_tel");

                textBox67.Clear();
                textBox67.DataBindings.Clear();
                textBox67.DataBindings.Add("text", dtst1, "firma.f_fax");

                textBox68.Clear();
                textBox68.DataBindings.Clear();
                textBox68.DataBindings.Add("text", dtst1, "firma.f_mail");

                textBox69.Clear();
                textBox69.DataBindings.Clear();
                textBox69.DataBindings.Add("text", dtst1, "firma.f_adres");


                textBox70.Clear();
                textBox70.DataBindings.Clear();
                textBox70.DataBindings.Add("text", dtst1, "firma.f_vergi_dairesi");

                textBox71.Clear();
                textBox71.DataBindings.Clear();
                textBox71.DataBindings.Add("text", dtst1, "firma.f_vergi_no");

                comboBox4.Text = "";
                comboBox4.DataBindings.Clear();
                comboBox4.DataBindings.Add("text", dtst1, "firma.f_hizmet");

                DG5.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(@textBox11.Text);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(dosya_yolu);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                dosya_sec(listBox1);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox13.Text != "")
                {
                    textBox29.Text = textBox13.Text;
                    conn.Close();

                    OleDbCommand com = new OleDbCommand("insert into web(w_f_adi,w_panel,w_mail,w_sifre,w_alan_adi,w_hosting,w_aciklama) values ('" + textBox13.Text + "','" + textBox14.Text + "','" + textBox15.Text + "','" + textBox16.Text + "','" + textBox17.Text + "','" + textBox18.Text + "','" + textBox19.Text + "')", conn);
                    MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    com.Connection.Open();
                    com.ExecuteNonQuery();
                    conn.Close();
                    dosya_ekle(textBox13, textBox20, listBox1);
                    button13.PerformClick();
                }
                else
                {
                    MessageBox.Show("Lütfen Bir Firma Adı Arattınız Ve Ya Listeden Seçiniz", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           

       }
        private void button12_Click(object sender, EventArgs e)
        {

            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_adi='" + textBox21.Text + "' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG2.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();

                web_show();
                firma_show(textBox21, textBox13, textBox13, textBox22, textBox23, textBox24, textBox25, textBox26, textBox27, textBox28, textBox29, comboBox1, textBox31);
                dosya_getir(listBox2, textBox31);

                if (textBox13.Text == "")
                    MessageBox.Show("Kayıt Bulunamadı. Lütfen Aradığınız Firma Adını Kontrol Ediniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }      
        private void tabPage6_Enter(object sender, EventArgs e)
        {
            try
            {
                web_getir();
                hizmet_temizle(groupBox11, groupBox7, groupBox8, groupBox9, groupBox12);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void tabPage2_Enter(object sender, EventArgs e)
        {
            try
            {

                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_hizmet='WEB' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG2.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();

                web_getir();
                hizmet_temizle(groupBox11, groupBox7, groupBox8, groupBox9, groupBox12);
                promosyon_getir();
                hizmet_temizle(groupBox14, groupBox17, groupBox18, groupBox16, groupBox13);
                reklam_getir();
                hizmet_temizle(groupBox20, groupBox23, groupBox24, groupBox22, groupBox19);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
        private void DG2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hizmet_temizle(groupBox11, groupBox7, groupBox8, groupBox9, groupBox12);
                textBox21.Clear();
                textBox13.Text = (DG2.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox22.Text = (DG2.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox23.Text = (DG2.Rows[e.RowIndex].Cells[3].Value.ToString());
                textBox24.Text = (DG2.Rows[e.RowIndex].Cells[4].Value.ToString());
                textBox25.Text = (DG2.Rows[e.RowIndex].Cells[5].Value.ToString());
                textBox26.Text = (DG2.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox27.Text = (DG2.Rows[e.RowIndex].Cells[7].Value.ToString());
                textBox28.Text = (DG2.Rows[e.RowIndex].Cells[8].Value.ToString());
                comboBox1.Text = (DG2.Rows[e.RowIndex].Cells[9].Value.ToString());
                textBox29.Text = (DG2.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox31.Text = (DG2.Rows[e.RowIndex].Cells[10].Value.ToString());

                web_kont();

                if (web == true)
                {
                    web_show();
                    dosya_getir(list1: listBox2, text: textBox31);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                web_getir();
                hizmet_temizle(groupBox11, groupBox7, groupBox8, groupBox9, groupBox12);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {

                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update web set w_f_adi='" + textBox29.Text + "',w_panel='" + textBox14.Text + "',w_mail='" + textBox15.Text + "',w_sifre='" + textBox16.Text + "',w_alan_adi='" + textBox17.Text + "',w_hosting='" + textBox18.Text + "',w_aciklama='" + textBox19.Text + "'where w_f_adi='" + textBox29.Text + "'", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();
                dosya_ekle(textBox13, textBox20, listBox1);
                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button13.PerformClick();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
      
       
        private void tabPage7_Enter(object sender, EventArgs e)
        {
            try
            {
                promosyon_getir();
                hizmet_temizle(groupBox14, groupBox17, groupBox18, groupBox16, groupBox13);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DG3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hizmet_temizle(groupBox14, groupBox17, groupBox18, groupBox16, groupBox13);
                textBox43.Clear();
                textBox42.Text = (DG3.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox35.Text = (DG3.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox36.Text = (DG3.Rows[e.RowIndex].Cells[3].Value.ToString());
                textBox37.Text = (DG3.Rows[e.RowIndex].Cells[4].Value.ToString());
                textBox38.Text = (DG3.Rows[e.RowIndex].Cells[5].Value.ToString());
                textBox39.Text = (DG3.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox40.Text = (DG3.Rows[e.RowIndex].Cells[7].Value.ToString());
                textBox41.Text = (DG3.Rows[e.RowIndex].Cells[8].Value.ToString());
                comboBox2.Text = (DG3.Rows[e.RowIndex].Cells[9].Value.ToString());
                textBox50.Text = (DG3.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox32.Text = (DG3.Rows[e.RowIndex].Cells[10].Value.ToString());

                promosyon_kont();

                if (promosyon == true)
                {
                    promosyon_show();
                    dosya_getir(listBox3, textBox32);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                hizmet_temizle(groupBox11, groupBox7, groupBox8, groupBox9, groupBox12);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                promosyon_getir();
                hizmet_temizle(groupBox14, groupBox17, groupBox18, groupBox16, groupBox13);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                dosya_sec(listBox4);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox42.Text != "")
                {
                    textBox50.Text = textBox42.Text;
                    conn.Close();

                    OleDbCommand com = new OleDbCommand("insert into promosyon(p_f_adi,p_urun_bilgi) values ('" + textBox42.Text + "','" + textBox44.Text + "')", conn);
                    MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    com.Connection.Open();
                    com.ExecuteNonQuery();
                    conn.Close();
                    dosya_ekle(textBox42, textBox34, listBox4);
                    button15.PerformClick();
                }
                else
                {
                    MessageBox.Show("Lütfen Bir Firma Adı Arattınız Ve Ya Listeden Seçiniz", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
               
            
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_adi='" + textBox43.Text + "' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG3.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();

                promosyon_show();
                firma_show(textBox43, textBox42, textBox42, textBox35, textBox36, textBox37, textBox38, textBox39, textBox40, textBox41, textBox50, comboBox2, textBox32);
                dosya_getir(listBox3, textBox32);

                if (textBox42.Text == "")
                    MessageBox.Show("Kayıt Bulunamadı. Lütfen Aradığınız Firma Adını Kontrol Ediniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update promosyon set p_f_adi='" + textBox50.Text + "',p_urun_bilgi='" + textBox44.Text + "' where p_f_adi='" + textBox50.Text + "' ", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();
                dosya_ekle(textBox42, textBox34, listBox4);
                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button15.PerformClick();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Bu İşlem Sadece Eklenen Web Bilgilerini Siler \n\nSilmek İstediğinizden Eminmisiniz", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    conn.Close();

                    OleDbCommand sil = new OleDbCommand("delete from web where w_f_adi= '" + textBox29.Text + "' ", conn);
                    sil.Connection.Open();
                    sil.ExecuteNonQuery();
                    conn.Close();
                    if (MessageBox.Show("Silme İşlemi Tamamlandı \n\nDosyaları Silmek İster Misiniz?", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        dosya_sil(textBox13, textBox20, listBox2);

                    }
                    button13.PerformClick();
                }
                else
                {
                    MessageBox.Show("Silme İşlemi İptal Edildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {

                if (MessageBox.Show("Bu İşlem Sadece Eklenen Promosyon Bilgisini Siler \n\nSilmek İstediğinizden Eminmisiniz", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    conn.Close();

                    OleDbCommand sil = new OleDbCommand("delete from promosyon where p_f_adi= '" + textBox50.Text + "' ", conn);
                    sil.Connection.Open();
                    sil.ExecuteNonQuery();
                    conn.Close();
                    if (MessageBox.Show("Silme İşlemi Tamamlandı \n\nDosyaları Silmek İster Misiniz?", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        dosya_sil(textBox42, textBox34, listBox3);


                    }
                    button15.PerformClick();
                }
                else
                {
                    MessageBox.Show("Silme İşlemi İptal Edildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From firma  where f_adi='" + textBox57.Text + "' ", conn);
                adtr1.Fill(dtst1, "firma");
                DG4.DataSource = dtst1.Tables["firma"];
                adtr1.Dispose();
                conn.Close();

                reklam_show();
                firma_show(textBox57, textBox56, textBox56, textBox48, textBox49, textBox51, textBox52, textBox53, textBox54, textBox55, textBox59, comboBox3, textBox45);
                dosya_getir(listBox5, textBox45);

                if (textBox56.Text == "")
                    MessageBox.Show("Kayıt Bulunamadı. Lütfen Aradığınız Firma Adını Kontrol Ediniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage8_Enter(object sender, EventArgs e)
        {
            try
            {
                reklam_getir();
                hizmet_temizle(groupBox20, groupBox23, groupBox24, groupBox22, groupBox19);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                reklam_getir();
                hizmet_temizle(groupBox20, groupBox23, groupBox24, groupBox22, groupBox19);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox56.Text != "")
                {
                    textBox59.Text = textBox56.Text;
                    conn.Close();

                    OleDbCommand com = new OleDbCommand("insert into reklam(r_f_adi,r_bilgi) values ('" + textBox59.Text + "','" + textBox58.Text + "')", conn);
                    MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    com.Connection.Open();
                    com.ExecuteNonQuery();
                    conn.Close();
                    dosya_ekle(textBox56, textBox47, listBox6);
                    button22.PerformClick();
                }
                else
                {
                    MessageBox.Show("Lütfen Bir Firma Adı Arattınız Ve Ya Listeden Seçiniz", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update reklam set r_f_adi='" + textBox59.Text + "',r_bilgi='" + textBox58.Text + "' where r_f_adi='" + textBox59.Text + "' ", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();
                dosya_ekle(textBox56, textBox47, listBox6);
                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button22.PerformClick();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            try
            {

                if (MessageBox.Show("Bu İşlem Sadece Eklenen Reklam Bilgisini Siler \n\nSilmek İstediğinizden Eminmisiniz", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    conn.Close();

                    OleDbCommand sil = new OleDbCommand("delete from reklam where r_f_adi= '" + textBox59.Text + "' ", conn);
                    sil.Connection.Open();
                    sil.ExecuteNonQuery();
                    conn.Close();
                    if (MessageBox.Show("Silme İşlemi Tamamlandı \n\nDosyaları Silmek İster Misiniz?", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                    {
                        dosya_sil(textBox56, textBox47, listBox5);


                    }
                    button22.PerformClick();
                }
                else
                {
                    MessageBox.Show("Silme İşlemi İptal Edildi", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try
            {
                dosya_sec(listBox6);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DG4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                hizmet_temizle(groupBox20, groupBox23, groupBox24, groupBox22, groupBox19);
                textBox57.Clear();
                textBox56.Text = (DG4.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox48.Text = (DG4.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox49.Text = (DG4.Rows[e.RowIndex].Cells[3].Value.ToString());
                textBox51.Text = (DG4.Rows[e.RowIndex].Cells[4].Value.ToString());
                textBox52.Text = (DG4.Rows[e.RowIndex].Cells[5].Value.ToString());
                textBox53.Text = (DG4.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox54.Text = (DG4.Rows[e.RowIndex].Cells[7].Value.ToString());
                textBox55.Text = (DG4.Rows[e.RowIndex].Cells[8].Value.ToString());
                comboBox3.Text = (DG4.Rows[e.RowIndex].Cells[9].Value.ToString());
                textBox59.Text = (DG4.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox45.Text = (DG4.Rows[e.RowIndex].Cells[10].Value.ToString());

                reklam_kont();

                if (reklam == true)
                {
                    reklam_show();
                    dosya_getir(listBox5, textBox45);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    

        private void DG5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                pictureBox1.Visible = false;
                richTextBox1.Visible = false;
                dataGridView2.Visible = false;
                webBrowser1.Visible = false;
                textBox72.Text = (DG5.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox65.Text = (DG5.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox66.Text = (DG5.Rows[e.RowIndex].Cells[3].Value.ToString());
                textBox67.Text = (DG5.Rows[e.RowIndex].Cells[4].Value.ToString());
                textBox68.Text = (DG5.Rows[e.RowIndex].Cells[5].Value.ToString());
                textBox69.Text = (DG5.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox70.Text = (DG5.Rows[e.RowIndex].Cells[7].Value.ToString());
                textBox71.Text = (DG5.Rows[e.RowIndex].Cells[8].Value.ToString());
                comboBox4.Text = (DG5.Rows[e.RowIndex].Cells[9].Value.ToString());
                textBox11.Text = (DG5.Rows[e.RowIndex].Cells[10].Value.ToString());
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(@textBox11.Text);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button20_Click_1(object sender, EventArgs e)
        {
            try
            {
                takvim.BringToFront();
                takvim.Visible = true;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void takvim_DateSelected(object sender, DateRangeEventArgs e)
        {
            try
            {
                string tarih = takvim.SelectionStart.Date.ToShortDateString();

                maskedTextBox2.Text = tarih;
                takvim.SendToBack();
                takvim.Visible = false;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                maskedTextBox1.Text = comboBox5.Text + comboBox6.Text;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                maskedTextBox1.Text = comboBox5.Text + comboBox6.Text;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    
        private void tabPage4_Enter(object sender, EventArgs e)
        {
            try
            {
                secimi_iptal(groupBox28);
                is_tablosu_getir(DG6, comboBox7);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                OleDbCommand com1 = new OleDbCommand("insert into isler(i_f_adi,i_hizmet,i_bilgi,i_tarih,i_saat,i_personel) values ('" + textBox60.Text + "','" + textBox62.Text + "','" + textBox61.Text + "','" + maskedTextBox2.Text + "','" + maskedTextBox1.Text + "','" + comboBox7.Text + "')", conn);
                com1.Connection.Open();
                com1.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                is_tablosu_getir(DG6, comboBox7);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox60_TextChanged(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select f_hizmet From firma  where f_adi like'%" + textBox60.Text + "%'", conn);
                adtr1.Fill(dtst1, "firma");
                textBox62.Text = "Hizmet";
                textBox62.DataBindings.Clear();
                textBox62.DataBindings.Add("text", dtst1, "firma.f_hizmet");

                adtr1.Dispose();
                conn.Close();

                DataSet dtst2 = new DataSet();
                OleDbDataAdapter adtr2 = new OleDbDataAdapter("select i_id From isler  where i_f_adi like'%" + textBox60.Text + "%'", conn);
                adtr2.Fill(dtst2, "isler");
                id.Clear();
                id.DataBindings.Clear();
                id.DataBindings.Add("text", dtst2, "isler.i_id");

                adtr1.Dispose();
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        
        private void DG6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox60.Enabled = false;
                button29.Enabled = false;

                textBox63.Text = (DG6.Rows[e.RowIndex].Cells[0].Value.ToString());
                textBox60.Text = (DG6.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox62.Text = (DG6.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox61.Text = (DG6.Rows[e.RowIndex].Cells[3].Value.ToString());
                maskedTextBox2.Text = (DG6.Rows[e.RowIndex].Cells[4].Value.ToString());
                maskedTextBox1.Text = (DG6.Rows[e.RowIndex].Cells[5].Value.ToString());
                comboBox7.Text = (DG6.Rows[e.RowIndex].Cells[6].Value.ToString());
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update isler set i_f_adi='" + textBox60.Text + "',i_hizmet='" + textBox62.Text + "',i_bilgi='" + textBox61.Text + "' ,i_tarih='" + maskedTextBox2.Text + "',i_saat='" + maskedTextBox1.Text + "',i_personel='" + comboBox7.Text + "' where i_f_adi='" + textBox60.Text + "'AND i_id=" + textBox63.Text + " ", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                secimi_iptal(groupBox28);
                is_tablosu_getir(DG6, comboBox7);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            try
            {
                is_tablosu_getir(DG6, comboBox7);
                secimi_iptal(groupBox28);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            try
            {
                maskedTextBox6.Text = DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString();
                //is_tablosu_getir(DG7, comboBox10);
                ise_gonder(DG7);
                is_temizle();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                timer1.Interval = 900000;
                panel1.Controls.Clear();

                y_ekseni = -10;
                string[] parca_saat;

                //maskedTextBox5.Text = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();

                saat_n.Text = DateTime.Now.Hour.ToString();
                dakika_n.Text = DateTime.Now.Minute.ToString();

                isler.Items.Clear();
                idler.Items.Clear();
                //comboBox14.Items.Clear();

                sa_n = Convert.ToInt32(saat_n.Text);
                dk_n = Convert.ToInt32(dakika_n.Text);

                int is_sayisi = 0;
                conn.Close();
                conn.Open();
                OleDbCommand com = new OleDbCommand("select i_saat,i_id from isler where i_durum='0' AND i_tarih='" + maskedTextBox6.Text + "'", conn);
                OleDbDataReader dr = com.ExecuteReader();

                while (dr.Read())
                {
                    isler.Items.Add(dr["i_saat"]);
                    idler.Items.Add(dr["i_id"]);
                    is_sayisi++;
                }


                for (int i = 0; i < is_sayisi; i++)
                {

                    isler.SelectedIndex = i;
                    idler.SelectedIndex = i;

                    saat.Text = isler.SelectedItem.ToString();
                    parca_saat = saat.Text.Split(':');//Belli Bir Karakterden Sonrasını Alma
                    saat_1.Text = parca_saat[0];
                    dakika.Text = parca_saat[1];

                    sa = Convert.ToInt32(saat_1.Text);
                    dk = Convert.ToInt32(dakika.Text);

                    if (dk < dk_n)
                    {
                        dk += 60;
                        sa -= 1;

                        dk_fark = dk - dk_n;
                        sa_fark = sa - sa_n;
                        sa_fark = sa_fark * 60;
                        dk_fark += sa_fark;

                    }
                    if (dk > dk_n)
                    {
                        dk_fark = dk - dk_n;
                        sa_fark = sa - sa_n;
                        sa_fark = sa_fark * 60;
                        dk_fark += sa_fark;
                    }
                    if (dk_fark <= 0)
                    {

                        y_ekseni += 20;

                        uyari u1 = new uyari(this);
                        u1.Location = new System.Drawing.Point(u1.Location.X, u1.Location.Y + y_ekseni);
                        u1.Text = "UYARI: " + idler.SelectedItem.ToString() + ". İş'in Zamanı Geçti ";

                        panel1.Controls.Add(u1);

                        (new SoundPlayer("blip.wav")).Play();
                        groupBox35.BringToFront();
                        groupBox35.Visible = true;

                    }
                    else if (dk_fark <= 30)
                    {

                        y_ekseni += 20;

                        uyari u1 = new uyari(this);
                        u1.Location = new System.Drawing.Point(u1.Location.X, u1.Location.Y + y_ekseni);
                        u1.Text = "UYARI: " + idler.SelectedItem.ToString() + ". İş İçin " + dk_fark.ToString() + " Dakika Zaman Kaldı";

                        panel1.Controls.Add(u1);

                        (new SoundPlayer("blip.wav")).Play();
                        groupBox35.BringToFront();
                        groupBox35.Visible = true;

                    }

                    comboBox14.Items.Add(i + 1 + ". işe" + dk_fark.ToString() + " dakika kaldı");
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
          
        }
      


        private void button36_Click_1(object sender, EventArgs e)
        {
            try
            {
                groupBox35.SendToBack();
                groupBox35.Visible = false;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst2 = new DataSet();
                OleDbDataAdapter adtr2 = new OleDbDataAdapter();
                if (maskedTextBox6.Text != null)
                {
                    if (textBox80.Text == "")
                    {
                        adtr2 = new OleDbDataAdapter("select * From isler  where i_f_adi='" + textBox75.Text + "' AND i_tarih='" + maskedTextBox6.Text + "' ", conn);
                    }
                    else
                    {
                        adtr2 = new OleDbDataAdapter("select * From isler  where i_f_adi='" + textBox80.Text + "'AND i_tarih='" + maskedTextBox6.Text + "' ", conn);
                    }
                }

                else
                {

                    if (textBox80.Text == "")
                    {
                        adtr2 = new OleDbDataAdapter("select * From isler  where i_f_adi='" + textBox75.Text + "'", conn);
                    }
                    else
                    {
                        adtr2 = new OleDbDataAdapter("select * From isler  where i_f_adi='" + textBox80.Text + "' ", conn);
                    }

                }
                adtr2.Fill(dtst2, "isler");

                textBox64.Clear();
                textBox64.DataBindings.Clear();
                textBox64.DataBindings.Add("text", dtst2, "isler.i_id");

                textBox75.Clear();
                textBox75.DataBindings.Clear();
                textBox75.DataBindings.Add("text", dtst2, "isler.i_f_adi");

                textBox73.Clear();
                textBox73.DataBindings.Clear();
                textBox73.DataBindings.Add("text", dtst2, "isler.i_hizmet");

                textBox74.Clear();
                textBox74.DataBindings.Clear();
                textBox74.DataBindings.Add("text", dtst2, "isler.i_bilgi");

                maskedTextBox3.Clear();
                maskedTextBox3.DataBindings.Clear();
                maskedTextBox3.DataBindings.Add("text", dtst2, "isler.i_tarih");

                maskedTextBox4.Clear();
                maskedTextBox4.DataBindings.Clear();
                maskedTextBox4.DataBindings.Add("text", dtst2, "isler.i_saat");

                comboBox10.Text = "PERSONEL SEÇİNİZ";
                comboBox10.DataBindings.Clear();
                comboBox10.DataBindings.Add("text", dtst2, "isler.i_personel");

                textBox82.Clear();
                textBox82.DataBindings.Clear();
                textBox82.DataBindings.Add("text", dtst2, "isler.i_durum");
                if (textBox82.Text == "1")
                {
                    textBox82.Text = "PERSONEL GÖNDERİLDİ";
                }
                else if (textBox82.Text == "0")
                {
                    textBox82.Text = "PERSONEL GÖNDERİLMEDİ";
                }



                ///PERSONEL BİLGİSİ ÇEKİLİYOR

                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter();

                adtr1 = new OleDbDataAdapter("select * From personel  where Person_Adi='" + comboBox10.Text + "' ", conn);

                adtr1.Fill(dtst1, "personel");

                textBox76.Clear();
                textBox76.DataBindings.Clear();
                textBox76.DataBindings.Add("text", dtst1, "personel.person_id");

                textBox77.Clear();
                textBox77.DataBindings.Clear();
                textBox77.DataBindings.Add("text", dtst1, "personel.Person_Adi");

                textBox78.Clear();
                textBox78.DataBindings.Clear();
                textBox78.DataBindings.Add("text", dtst1, "personel.Person_Tel");

                textBox79.Clear();
                textBox79.DataBindings.Clear();
                textBox79.DataBindings.Add("text", dtst1, "personel.Person_Mail");
                firma_show(textBox80, textBox75, textBox84, textBox85, textBox86, textBox87, textBox88, textBox89, textBox90, textBox91, textBox75, comboBox8, textBox81);

            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        
        }

        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                takvim1.BringToFront();
                takvim1.Visible = true;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void takvim1_DateSelected(object sender, DateRangeEventArgs e)
        {
            try
            {
                string tarih1 = takvim1.SelectionStart.Date.ToShortDateString();

                maskedTextBox6.Text = tarih1;
                takvim1.SendToBack();
                takvim1.Visible = false;
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       
        private void button34_Click(object sender, EventArgs e)
        {
            try
            {
                texte_yaz();
                MessageBox.Show(richTextBox2.Text, "İş Bilgileri", MessageBoxButtons.OK, MessageBoxIcon.Information);
                int i = 0;
                while (i == 0)
                {
                    if ((MessageBox.Show("Çıktı Alamak İstermisiniz", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
                    {
                        i = 1;
                        button33.PerformClick();
                    }
                    else
                    {
                        if ((MessageBox.Show("İş Bilgilerini Tekrar Yazdıramayacaksınız Devem Etmek İster Misiniz?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
                        {
                            i = 1;

                        }
                        else
                        {
                            i = 0;
                        }
                    }
                }

                conn.Close();
                OleDbCommand gncl = new OleDbCommand("update isler set i_durum='1' where i_f_adi='" + textBox75.Text + "'AND i_id=" + textBox64.Text + " ", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();
                is_temizle();
                MessageBox.Show("İş Atama İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DG7_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {

                textBox64.Text = (DG7.Rows[e.RowIndex].Cells[0].Value.ToString());
                textBox75.Text = (DG7.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox73.Text = (DG7.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox74.Text = (DG7.Rows[e.RowIndex].Cells[3].Value.ToString());
                maskedTextBox3.Text = (DG7.Rows[e.RowIndex].Cells[4].Value.ToString());
                maskedTextBox4.Text = (DG7.Rows[e.RowIndex].Cells[5].Value.ToString());
                comboBox10.Text = (DG7.Rows[e.RowIndex].Cells[6].Value.ToString());
                textBox82.Text = (DG7.Rows[e.RowIndex].Cells[7].Value.ToString());
                if (textBox82.Text == "1")
                {
                    textBox82.Text = "PERSONEL GÖNDERİLDİ";
                }
                else if (textBox82.Text == "0")
                {
                    textBox82.Text = "PERSONEL GÖNDERİLMEDİ";
                }

                ///PERSONEL BİLGİSİ ÇEKİLİYOR

                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter();

                adtr1 = new OleDbDataAdapter("select * From personel  where Person_Adi='" + comboBox10.Text + "' ", conn);

                adtr1.Fill(dtst1, "personel");

                textBox76.Clear();
                textBox76.DataBindings.Clear();
                textBox76.DataBindings.Add("text", dtst1, "personel.person_id");

                textBox77.Clear();
                textBox77.DataBindings.Clear();
                textBox77.DataBindings.Add("text", dtst1, "personel.Person_Adi");

                textBox78.Clear();
                textBox78.DataBindings.Clear();
                textBox78.DataBindings.Add("text", dtst1, "personel.Person_Tel");

                textBox79.Clear();
                textBox79.DataBindings.Clear();
                textBox79.DataBindings.Add("text", dtst1, "personel.Person_Mail");
                firma_show(textBox80, textBox75, textBox84, textBox85, textBox86, textBox87, textBox88, textBox89, textBox90, textBox91, textBox75, comboBox8, textBox81);

                conn.Close();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
          

        }

        private void button31_Click(object sender, EventArgs e)
        {
            try
            {
                is_temizle();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            try
            {
                texte_yaz();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                printDialog1.Document = printDocument1;

                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    StringReader reader = new StringReader(richTextBox2.Text);
                    printDocument1.Print();

                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
           
   
        private void DocumentToPrint_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                StringReader reader = new StringReader(richTextBox2.Text);
                float LinesPerPage = 0;
                float YPosition = 0;
                int Count = 0;
                float LeftMargin = e.MarginBounds.Left;
                float TopMargin = e.MarginBounds.Top;
                string Line = null;
                System.Drawing.Font PrintFont = this.richTextBox2.Font;
                SolidBrush PrintBrush = new SolidBrush(Color.Black);

                LinesPerPage = e.MarginBounds.Height / PrintFont.GetHeight(e.Graphics);

                while (Count < LinesPerPage && ((Line = reader.ReadLine()) != null))
                {
                    YPosition = TopMargin + (Count * PrintFont.GetHeight(e.Graphics));
                    e.Graphics.DrawString(Line, PrintFont, PrintBrush, LeftMargin, YPosition, new StringFormat());
                    Count++;
                }

                if (Line != null)
                {
                    e.HasMorePages = true;
                }
                else
                {
                    e.HasMorePages = false;
                }
                PrintBrush.Dispose();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void kullanici_getir()
        {
            try
            {
                textBox95.Enabled = true;
                button47.Enabled = true;

                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From kullanici ", conn);
                adtr1.Fill(dtst1, "kullanici");
                DG8.DataSource = dtst1.Tables["kullanici"];
                adtr1.Dispose();
                conn.Close();
                textBox93.Clear();
                textBox95.Clear();
                textBox96.Clear();
                comboBox9.Text = "Rütbe Seçiniz";
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void personel_getir()
        {
            try
            {
                textBox98.Enabled = true;
                button50.Enabled = true;
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From personel ", conn);
                adtr1.Fill(dtst1, "personel");
                DG9.DataSource = dtst1.Tables["personel"];
                adtr1.Dispose();
                conn.Close();
                textBox97.Clear();
                textBox98.Clear();
                textBox99.Clear();
                textBox100.Clear();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage9_Enter(object sender, EventArgs e)
        {
            try
            {
                tabControl3.Visible = false;
                panel3.Visible = true;
                kullanici_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                OleDbCommand com1 = new OleDbCommand("insert into kullanici(k_adi,k_sifre,k_rutbe) values ('" + textBox95.Text + "','" + textBox96.Text + "','" + comboBox9.Text + "')", conn);
                com1.Connection.Open();
                com1.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                kullanici_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From kullanici where k_adi='" + textBox93.Text + "'", conn);
                adtr1.Fill(dtst1, "kullanici");

                textBox95.Clear();
                textBox95.DataBindings.Clear();
                textBox95.DataBindings.Add("text", dtst1, "kullanici.k_adi");

                textBox96.Clear();
                textBox96.DataBindings.Clear();
                textBox96.DataBindings.Add("text", dtst1, "kullanici.k_sifre");

                comboBox9.Text = "Rütbe Seçiniz";
                comboBox9.DataBindings.Clear();
                comboBox9.DataBindings.Add("text", dtst1, "kullanici.k_rutbe");
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox95.Enabled = false;
                button47.Enabled = false;
                textBox95.Text = (DG8.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox96.Text = (DG8.Rows[e.RowIndex].Cells[2].Value.ToString());
                comboBox9.Text = (DG8.Rows[e.RowIndex].Cells[3].Value.ToString());
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button44_Click(object sender, EventArgs e)
        {
            try
            {
                button47.Enabled = true;
                textBox93.Clear();
                textBox95.Clear();
                textBox96.Clear();
                comboBox9.Text = "Rütbe Seçiniz";
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button46_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update kullanici set k_adi='" + textBox95.Text + "',k_sifre='" + textBox96.Text + "',k_rutbe='" + comboBox9.Text + "'where k_adi='" + textBox95.Text + "'", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();

                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                kullanici_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button51_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From personel where Person_Adi='" + textBox97.Text + "'", conn);
                adtr1.Fill(dtst1, "personel");

                textBox98.Clear();
                textBox98.DataBindings.Clear();
                textBox98.DataBindings.Add("text", dtst1, "personel.Person_Adi");

                textBox99.Clear();
                textBox99.DataBindings.Clear();
                textBox99.DataBindings.Add("text", dtst1, "personel.Person_Tel");

                textBox100.Clear();
                textBox100.DataBindings.Clear();
                textBox100.DataBindings.Add("text", dtst1, "personel.Person_Mail");
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                OleDbCommand com1 = new OleDbCommand("insert into personel(Person_Adi,Person_Tel,Person_Mail) values ('" + textBox98.Text + "','" + textBox99.Text + "','" + textBox100.Text + "')", conn);
                com1.Connection.Open();
                com1.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Kayıt Tamamlandı", "Kayıt İşlemi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                personel_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DG9_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                textBox98.Enabled = false;
                button50.Enabled = false;
                textBox98.Text = (DG9.Rows[e.RowIndex].Cells[1].Value.ToString());
                textBox99.Text = (DG9.Rows[e.RowIndex].Cells[2].Value.ToString());
                textBox100.Text = (DG9.Rows[e.RowIndex].Cells[3].Value.ToString());
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button52_Click(object sender, EventArgs e)
        {
            try
            {
                button50.Enabled = true;
                textBox97.Clear();
                textBox98.Clear();
                textBox99.Clear();
                textBox100.Clear();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage12_Enter(object sender, EventArgs e)
        {
            try
            {
                personel_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button49_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand gncl = new OleDbCommand("update personel set Person_Adi='" + textBox98.Text + "',Person_Tel='" + textBox99.Text + "',Person_Mail='" + textBox100.Text + "'where Person_Adi='" + textBox98.Text + "'", conn);
                gncl.Connection.Open();
                gncl.ExecuteNonQuery();
                conn.Close();

                MessageBox.Show("Güncelleme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                personel_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

   

        private void button48_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand sil = new OleDbCommand("delete from personel where Person_Adi='" + textBox98.Text + "' ", conn);
                sil.Connection.Open();
                sil.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Silme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                personel_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();

                OleDbCommand sil = new OleDbCommand("delete from kullanici where k_adi='" + textBox95.Text + "' ", conn);
                sil.Connection.Open();
                sil.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Silme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                kullanici_getir();
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                OleDbCommand sil = new OleDbCommand("delete from isler where i_f_adi='" + textBox60.Text + "' AND i_id=" + textBox63.Text + " ", conn);
                sil.Connection.Open();
                sil.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Silme İşlemi Tamamlandı", "Bilgi", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                secimi_iptal(groupBox28);
                is_tablosu_getir(DG6, comboBox7);
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button53_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From kullanici where k_adi='" + textBox105.Text + "'", conn);
                adtr1.Fill(dtst1, "kullanici");

                textBox103.Clear();
                textBox103.DataBindings.Clear();
                textBox103.DataBindings.Add("text", dtst1, "kullanici.k_adi");

                textBox102.Clear();
                textBox102.DataBindings.Clear();
                textBox102.DataBindings.Add("text", dtst1, "kullanici.k_sifre");

                textBox101.Clear();
                textBox101.DataBindings.Clear();
                textBox101.DataBindings.Add("text", dtst1, "kullanici.k_rutbe");


                if (textBox105.Text == textBox103.Text & textBox104.Text == textBox102.Text & textBox101.Text == "ADMİN")
                {

                    MessageBox.Show("Başarılı Giriş Yapıldı", "Hoşgeldiniz", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tabControl3.Visible = true;
                    textBox105.Clear();
                    textBox104.Clear();
                    textBox102.Clear();
                    textBox101.Clear();
                    textBox103.Clear();
                    panel3.Visible = false;
                }
                else if (textBox105.Text == "" || textBox104.Text == "")
                {
                    MessageBox.Show("Lütfen Gerekli Alanları Doldurun", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox103.Text == "")
                {
                    MessageBox.Show("Kullanıcı Adınız Hatalı", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox104.Text != textBox102.Text)
                {
                    MessageBox.Show("Şifreniz Hatalı", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox101.Text != "ADMİN")
                {
                    MessageBox.Show("Giriş İçin Uygun Kullanıcı Değilsiniz", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button54_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Close();
                conn.Open();
                DataSet dtst1 = new DataSet();
                OleDbDataAdapter adtr1 = new OleDbDataAdapter("select * From kullanici where k_adi='" + textBox107.Text + "'", conn);
                adtr1.Fill(dtst1, "kullanici");

                textBox110.Clear();
                textBox110.DataBindings.Clear();
                textBox110.DataBindings.Add("text", dtst1, "kullanici.k_adi");

                textBox109.Clear();
                textBox109.DataBindings.Clear();
                textBox109.DataBindings.Add("text", dtst1, "kullanici.k_sifre");

                textBox108.Clear();
                textBox108.DataBindings.Clear();
                textBox108.DataBindings.Add("text", dtst1, "kullanici.k_rutbe");


                if (textBox107.Text == textBox110.Text & textBox106.Text == textBox109.Text & (textBox108.Text == "ADMİN" || textBox108.Text == "KULLANICI"))
                {

                    MessageBox.Show("Başarılı Giriş Yapıldı", "Hoşgeldiniz", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tabControl1.Visible = true;
                    textBox106.Clear();
                    textBox107.Clear();
                    textBox108.Clear();
                    textBox109.Clear();
                    textBox110.Clear();
                    panel4.Visible = false;
                    timer1.Enabled = true;

                }
                else if (textBox107.Text == "" || textBox106.Text == "")
                {
                    MessageBox.Show("Lütfen Gerekli Alanları Doldurun", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox110.Text == "")
                {
                    MessageBox.Show("Kullanıcı Adınız Hatalı", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox106.Text != textBox109.Text)
                {
                    MessageBox.Show("Şifreniz Hatalı", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                else if (textBox108.Text != "ADMİN" || textBox108.Text != "KULLANICI")
                {
                    MessageBox.Show("Giriş İçin Uygun Kullanıcı Değilsiniz", "HATA", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }

            }
            catch
            {
                MessageBox.Show("Bir Hata Oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    
     }
}
 

