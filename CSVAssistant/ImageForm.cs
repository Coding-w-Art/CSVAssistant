using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSVAssistant
{
    public partial class ImageForm : Form
    {
        bool mouseDown = false;
        bool mouseMove = false;
        Point mPoint = new Point();

        public ImageForm(string path)
        {
            InitializeComponent();
            SetImage(path);
            pictureBox1.MouseWheel += PictureBox1_MouseWheel;
        }

        private void ImageForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape || e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space)
            {
                Hide();
            }
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            if (!mouseMove)
                Hide();
        }

        public void SetImage(string path)
        {
            pictureBox1.ImageLocation = path;
        }

        private void PictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                mPoint.X = e.X;
                mPoint.Y = e.Y;
                mouseDown = true;
            }
        }

        private void PictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            mouseMove = false;
            mouseDown = false;
        }

        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                mouseMove = true;
                Point newPosition = MousePosition;
                newPosition.Offset(-mPoint.X, -mPoint.Y);
                Location = newPosition;

            }
        }

        private void PictureBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            SizeF size = new SizeF();
            int delta = e.Delta;
            if (delta < 0)
            {
                if (Size.Width <= 32 || Size.Height <= 32)
                    return;
                size.Width = 0.8f;
                size.Height = 0.8f;
            }
            else
            {
                if (Size.Width >= 640 || Size.Height >= 640)
                    return;

                size.Width = 1.25f;
                size.Height = 1.25f;
            }
            pictureBox1.Scale(size);
        }
    }
}
