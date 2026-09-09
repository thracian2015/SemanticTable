using System;
using System.Drawing;
using System.Windows.Forms;

namespace SemanticTable
{
    internal static class ScaledIcons
    {
        // Match the 96-DPI, 13-pixel font baseline used by the control layouts.
        internal static int SizeFor(Control control) => Math.Min(256,
            Math.Max(16, (int)Math.Round(16 * Math.Max(control.DeviceDpi / 96F, control.Font.Height / 13F))));

        internal static void Bind(ButtonBase button, Func<int, Bitmap> draw)
        {
            Action update = () =>
            {
                var size = SizeFor(button);
                if (button.Image != null && button.Image.Width == size) return;
                var previous = button.Image;
                button.Image = draw(size);
                previous?.Dispose();
            };
            button.FontChanged += (_, __) => update();
            button.DpiChangedAfterParent += (_, __) => update();
            button.HandleCreated += (_, __) => update();
            button.Disposed += (_, __) =>
            {
                var previous = button.Image;
                button.Image = null;
                previous?.Dispose();
            };
            update();
        }
    }
}
