using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelFileReformatter
{
    public partial class SetRowSkipRange : Form
    {
        private readonly ISetRowSkipRange Item;

        private SetRowSkipRange()
        {
            InitializeComponent();
        }

        public SetRowSkipRange(ISetRowSkipRange item) : this()
        {
            Item = item;

            if (item.RowSkipRange != null)
            {
                numericStart.Text = item.RowSkipRange.Item1.ToString();
                numericEnd.Text = item.RowSkipRange.Item2.ToString();
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var start = int.Parse(numericStart.Text);
            var end = int.Parse(numericEnd.Text);

            if (start > end)
            {
                MessageBox.Show("Range end must be greater than or equal to start", "Range Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Item.SetRowSkipRange(start, end);
        }
    }

    public interface ISetRowSkipRange
    {
        Tuple<int,int> RowSkipRange { get; }

        void SetRowSkipRange(int start, int end);
    }
}
