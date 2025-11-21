using System.Drawing;
using System.Windows.Forms;

public static class CheckedListBoxExtensions
{
    /// <summary>
    /// Ajusta el ancho del CheckedListBox en función del texto más largo.
    /// </summary>
    public static void AutoAdjustWidth(this CheckedListBox clb)
    {
        if (clb.Items.Count == 0)
            return;

        int maxWidth = 0;

        using (var g = clb.CreateGraphics())
        {
            foreach (var item in clb.Items)
            {
                string texto = clb.GetItemText(item);
                int ancho = (int)g.MeasureString(texto, clb.Font).Width;

                if (ancho > maxWidth)
                    maxWidth = ancho;
            }
        }

        // Espacio adicional para:
        // - Checkbox
        // - Scroll vertical (si llega a aparecer)
        // - Margen visual
        int padding = SystemInformation.VerticalScrollBarWidth + 25;

        clb.Width = maxWidth + padding;
    }
}

