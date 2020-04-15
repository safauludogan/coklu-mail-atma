using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Data.OleDb;

namespace coklu_mail_atma
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection baglan = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='mail_atma.accdb'");
        OleDbCommand sorgu = new OleDbCommand();
        OleDbDataReader dr;
        DataSet al = new DataSet();
        DataView goster = new DataView();
        OleDbDataAdapter verial;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                sorgu.Connection = baglan;
                sorgu.CommandText = "insert into iletisimbilgi(emailid,emailadresi,tckimlikno) values(" + int.Parse(textBox1.Text) + ",'" + textBox2.Text + "','" + textBox3.Text + "')";
                textBox1.Text = Convert.ToString(Convert.ToInt32(textBox1.Text) + 1);
                sorgu.ExecuteNonQuery();
                baglan.Close();
                MessageBox.Show("Kaydedildi");
                VeriTabaniniCek();
                baglan.Close();
                textBox2.Clear();
                textBox3.Clear();
            }
            catch
            {
                MessageBox.Show("Yanlış Bir İşlem Yaptınız");
            }
        }
        void VeriTabaniniCek()
        {
            listBox1.Items.Clear();
            al.Clear();
            verial = new OleDbDataAdapter("select * from iletisimbilgi", baglan);
            verial.Fill(al, "kitaplik");
            goster.Table = al.Tables[0];
            dataGridView1.DataSource = goster;

            //
            baglan.Close();
            baglan.Open();
            sorgu.Connection = baglan;
            sorgu.CommandText = "select * from iletisimbilgi";
            dr = sorgu.ExecuteReader();
            while (dr.Read())
            {
                listBox1.Items.Add(dr[1]);
            }
        }
        void EnSonKayit()
        {
            baglan.Close();
            baglan.Open();
            sorgu.Connection = baglan;
            sorgu.CommandText = "select max(emailid) from iletisimbilgi";
            dr = sorgu.ExecuteReader();
            dr.Read();
            try
            {
                textBox1.Text = Convert.ToString(Convert.ToInt32(dr[0]) + 1);
            }
            catch (InvalidCastException)
            {
                textBox1.Text = "1";
            }
            finally
            {
                baglan.Close();
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Maximum = 100;
            VeriTabaniniCek();
            textBox1.Enabled = false;
            EnSonKayit();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Start();
            progressBar1.Value = 0;
        }
        int anahtar = 0, kilit = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (anahtar == 0)
            {
                string[] dizi = new string[listBox1.Items.Count];               
            }
            anahtar = 1;
            progressBar1.Maximum = (listBox1.Items.Count * listBox1.Items.Count);
            progressBar1.Minimum = 0;
            if (progressBar1.Value <= (listBox1.Items.Count * listBox1.Items.Count))
            {
                try
                {
                    string[] dizi = new string[listBox1.Items.Count];
                    for (int i = 0; i < listBox1.Items.Count; i++)
                    {
                        if (kilit == 0)
                        {

                            progressBar1.Value += listBox1.Items.Count;
                            button2.Enabled = false;
                            button1.Enabled = false;
                            textBox1.Enabled = false;
                            textBox2.Enabled = false;
                            textBox3.Enabled = false;
                            textBox4.Enabled = false;
                            textBox5.Enabled = false;
                            listBox1.Enabled = false;
                            dataGridView1.Enabled = false;
                            dizi[i] = listBox1.Items[i].ToString();
                            SmtpClient sc = new SmtpClient();
                            sc.Port = 587;
                            sc.Host = "smtp.gmail.com";
                            sc.EnableSsl = true;
                            string konu = textBox4.Text;
                            string icerik = textBox5.Text;
                            sc.Credentials = new NetworkCredential("seyitgerizli@gmail.com", "senveben1996");
                            MailMessage mail = new MailMessage();
                            mail.From = new MailAddress("seyitgerizli@gmail.com", "Seyit GERİZLİ");
                            mail.To.Add(dizi[i]);
                            mail.Subject = konu;
                            mail.IsBodyHtml = true;
                            mail.Body = icerik;
                            sc.Send(mail);
                        }
                    }
                }
                catch (Exception)
                {
                    progressBar1.Value = 0;
                    kilit = 1;
                    timer1.Stop();
                    MessageBox.Show("Bir Hata İle Karşılaşıldı , yanlış bir mail girmiş olabilirsiniz veya internet bağlantınızı kontrol ediniz");
                    Form1 frm = new Form1();
                    frm.Show();
                    this.Hide();
                   
                }
                finally
                {
                    if (kilit == 0) 
                    {
                        timer1.Stop();
                        MessageBox.Show("Mailler Atıldı");
                        progressBar1.Value = 0;
                        anahtar = 0;
                        Form1 frm = new Form1();
                        frm.Show();
                        this.Hide();  
                    }                    
                }
            }
            else
            {
                anahtar = 0;
                timer1.Stop();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
        int s;
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            s = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells["emailid"].Value);
            try
            {
                baglan.Open();
                sorgu.Connection = baglan;
                sorgu.CommandText = "select * from iletisimbilgi where emailid=" + s + "";
                dr = sorgu.ExecuteReader();
                dr.Read();
                textBox1.Text = dr[0].ToString();
                textBox2.Text = dr[1].ToString();
                textBox3.Text = dr[2].ToString();
                baglan.Close();
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("Kayıt Bulunamadı");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
            }
            finally
            {
                baglan.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                sorgu.Connection = baglan;
                sorgu.CommandText = "update iletisimbilgi set emailadresi='" + textBox2.Text.Trim() + "',tckimlikno='" + textBox3.Text.Trim() + "' where  emailid=" + int.Parse(textBox1.Text) + "";
                sorgu.ExecuteNonQuery();
                VeriTabaniniCek();
                MessageBox.Show("İşlem başarı ile tamamlandı");
            }
            catch 
            {
                MessageBox.Show("Yanlış bir işlem yaptınız");
            }
            finally
            {
                EnSonKayit();
                textBox2.Clear();
                textBox3.Clear();
                baglan.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                baglan.Open();
                sorgu.Connection = baglan;
                sorgu.CommandText = "delete * from iletisimbilgi where emailid=" + int.Parse(textBox1.Text) + "";
                sorgu.ExecuteNonQuery();
                MessageBox.Show("Kayıt Silindi");
            }
            catch
            {
                MessageBox.Show("Kayıt Bulunamadı");
            }
            finally
            {
                VeriTabaniniCek();
                EnSonKayit();
                textBox3.Clear();
                textBox2.Clear();
                baglan.Close();
            }
        }
    }
}