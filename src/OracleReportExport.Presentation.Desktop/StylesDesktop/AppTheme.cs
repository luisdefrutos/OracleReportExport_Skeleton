using System.Drawing;

namespace OracleReportExport.Presentation.Desktop
{
    /// <summary>
    /// Tema visual centralizado de la aplicación.
    /// Modificar colores aquí cambia el aspecto completo.
    /// </summary>
    public class AppTheme
    {
        // --- Fondos generales ---
        public static readonly Color FormBackColor = Color.FromArgb(245, 247, 250);
        public static readonly Color TopPanelBackColor = Color.FromArgb(235, 239, 245);
        public static readonly Color GroupBoxBackColor = Color.FromArgb(245, 247, 250);

        // --- Texto ---
        public static readonly Color PrimaryTextColor = Color.FromArgb(40, 40, 40);
        public static readonly Color SecondaryTextColor = Color.FromArgb(90, 90, 90);
        public static readonly Color DisabledForeColor = Color.FromArgb(150, 150, 150);

        // --- Bordes generales ---
        public static readonly Color BorderColor = Color.FromArgb(200, 205, 210);

        // --- Acentos ---
        public static readonly Color AccentColor = Color.FromArgb(52, 120, 212);
        public static readonly Color AccentHoverColor = Color.FromArgb(42, 100, 184);
        public static readonly Color AccentSoftColor = Color.FromArgb(222, 231, 246);

        // --- Grids ---
        public static readonly Color GridBackColor = Color.White;
        public static readonly Color GridForeColor = PrimaryTextColor;
        public static readonly Color GridBorderColor = Color.FromArgb(210, 214, 220);

        // FILAS ALTERNAS (te faltaba este)
        public static readonly Color GridAlternateRowColor = Color.FromArgb(248, 250, 252);

        // Encabezados de columnas
        public static readonly Color GridHeaderBackColor = Color.FromArgb(240, 243, 248);
        public static readonly Color GridHeaderForeColor = PrimaryTextColor;

        // --- Botones ---
        public static readonly Color PrimaryButtonBackColor = AccentColor;
        public static readonly Color PrimaryButtonForeColor = Color.White;
        public static readonly Color PrimaryButtonBorderColor = AccentColor;

        public static readonly Color SecondaryButtonBackColor = Color.FromArgb(244, 246, 249);
        public static readonly Color SecondaryButtonForeColor = PrimaryTextColor;
        public static readonly Color SecondaryButtonBorderColor = BorderColor;

        // --- TabControl ---
        public static readonly Color ActiveTabBackColor = Color.White;
        public static readonly Color ActiveTabForeColor = PrimaryTextColor;

        public static readonly Color InactiveTabBackColor = TopPanelBackColor;
        public static readonly Color InactiveTabForeColor = SecondaryTextColor;

        public static readonly Color TabBorderColor = BorderColor;

        //  NUEVOS: bordes diferenciados en pestañas
        public static readonly Color ActiveTabBorderColor = AccentColor;
        public static readonly Color InactiveTabBorderColor = BorderColor;
    }
}
