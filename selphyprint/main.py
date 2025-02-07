import argparse
import os
from PIL import Image, UnidentifiedImageError
import numpy as np


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

fill_color = {
    "1": (0,),
    "L": (255,),
    "P": (0,),
    "RGB": (255, 255, 255),
    "RGBA": (255, 255, 255, 255),
    "CMYK": (0, 0, 0, 0),
    "YCbCr": (0, 0, 0, 0),
    "LAB": (100, 0, 0),
    "HSV": (0, 0, 255),
    "I": (2 ** 15,),
    "F": (1.0,),
    "I;16": (65535,),
    "I;16L": (65535,),
    "I;16B": (65535,),
    "I;16N": (65535,)
}

def convert_and_save(im, fn, dpi):
    base, ext = os.path.splitext(fn)
    out = im
    if im.mode not in ("RGB", "L"):
        if "I;16" in im.mode:
            x = np.floor(np.asarray(im)/256)
            out = Image.fromarray(x.astype(np.uint8))
        if "I" == im.mode:
            x = np.floor((np.asarray(im) + 2**15 - 1) / 256)
            out = Image.fromarray(x.astype(np.uint8))
        if "F" == im.mode:
            x = 255 * np.asarray(im)
            out = Image.fromarray(x.astype(np.uint8))
    out.save(fn, dpi=dpi)

def process_image(input_filename, border_pixels, output_filename):
    try:
        with Image.open(input_filename, "r") as im:
            c = fill_color[im.mode]
            background = Image.new(mode=im.mode, size=(WIDTH_PX, HEIGHT_PX), color=c)
            if im.height > im.width:
                im = im.rotate(angle=270, expand=True)
            scale_x = W / im.width
            scale_y = H / im.height
            scale = min(scale_x, scale_y)
            w = int(scale * im.width - border_pixels)
            h = int(scale * im.height - border_pixels)
            im = im.resize(size=(w, h), resample=Image.Resampling.NEAREST)
            offset_x = int((WIDTH_PX - w) / 2)
            offset_y = int((HEIGHT_PX - h) / 2)
            background.paste(im, box=(offset_x, offset_y))
            convert_and_save(im, output_filename, dpi=(DPI, DPI))
    except UnidentifiedImageError:
        print(f"Unsupported image file \"{input_filename}\"")


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
