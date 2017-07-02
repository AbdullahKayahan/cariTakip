using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
namespace Cari
{
    class uyari:Label
    {
         static Random r = new Random();
        Form1 f;
       
        public uyari(Form1 f)
        {
          
            this.f = f;
           
            this.Width = f.panel1.Width-20;
            this.Height = 20;

            this.ForeColor = System.Drawing.Color.White;
            this.Font = f.label1.Font;
            this.Visible = true;
           

        }  
    }
}
