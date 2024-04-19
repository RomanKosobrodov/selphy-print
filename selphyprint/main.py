import argparse
import os
from PIL import Image, UnidentifiedImageError

WIDTH = 150
HEIGHT = 100
INCH = 25.4
DPI = 300

SPACING = 1
STEPS = 10

LEFT = 7 * SPACING
RIGHT = 6 * SPACING
TOP = 4 * SPACING
BOTTOM = 4 * SPACING

WIDTH_PX = int(WIDTH / INCH * DPI)
HEIGHT_PX = int(HEIGHT / INCH * DPI)

X0 = int(LEFT / INCH * DPI)
X1 = int(RIGHT / INCH * DPI)
Y0 = int(TOP / INCH * DPI)
Y1 = int(BOTTOM / INCH * DPI)
W = WIDTH_PX - X1 - X0
H = HEIGHT_PX - Y1 - Y0


def process_image(input_filename, border_pixels, output_filename):
    background = Image.new(mode="RGB", size=(WIDTH_PX, HEIGHT_PX), color=(255, 255, 255))
    try:
        with Image.open(input_filename, "r") as im:
            if im.height > im.width:
                im = im.rotate(angle=90)
            scale_x = W / im.width
            scale_y = H / im.height
            scale = min(scale_x, scale_y)
            w = int(scale * im.width - border_pixels)
            h = int(scale * im.height - border_pixels)
            im = im.resize(size=(w, h), resample=Image.Resampling.NEAREST)
            offset_x = int((WIDTH_PX - w) / 2)
            offset_y = int((HEIGHT_PX - h) / 2)
            background.paste(im, box=(offset_x, offset_y))
            background.save(output_filename, dpi=(DPI, DPI))
    except UnidentifiedImageError:
        print(f"Unsupported image file \"{args.input}\"")


def main():
    parser = argparse.ArgumentParser(description="Adjust image to print on Canon Selphy without cropping")
    parser.add_argument("--input", "-i",
                        help="Input image or directory with multiple images",
                        type=str,
                        required=True)
    parser.add_argument("--border", "-b",
                        help="(optional) add border around the image (in mm)",
                        type=float,
                        default=0.0,
                        required=False)
    parser.add_argument("--output", "-o",
                        help="Output image or directory",
                        type=str,
                        required=True)
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Input path \"{args.input}\" does not exist")
        exit(1)

    output_parent = os.path.dirname(args.output)
    if not os.path.isdir(output_parent):
        print(f"Parent directory \"{output_parent}\" must exist")
        exit(2)

    border_px = int(args.border / INCH * DPI)

    if os.path.isfile(args.input):
        process_image(args.input, border_px, args.output)
    else:
        root, subdirectory, files = next(os.walk(args.input))
        for filename in files:
            fn = os.path.join(root, filename)
            basename, ext = os.path.splitext(filename)
            fn_output = os.path.join(args.output, basename + "-print" + ext)
            process_image(fn, border_px, fn_output)


if __name__ == "__main__":
    main()
