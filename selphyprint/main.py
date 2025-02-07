import argparse
import os
from PIL import Image, UnidentifiedImageError
from selphyprint.printing import get_printer_name, print_image


WIDTH = 150
HEIGHT = 100
INCH = 25.4
DPI = 300

SPACING = 1
STEPS = 10

LEFT = 7 * SPACING
RIGHT = 7 * SPACING
TOP = 4 * SPACING
BOTTOM = 4 * SPACING

WIDTH_PX = int(WIDTH / INCH * DPI)
HEIGHT_PX = int(HEIGHT / INCH * DPI)

MAX_BORDER = 30

X0 = int(LEFT / INCH * DPI)
X1 = int(RIGHT / INCH * DPI)
Y0 = int(TOP / INCH * DPI)
Y1 = int(BOTTOM / INCH * DPI)
W = WIDTH_PX - X1 - X0
H = HEIGHT_PX - Y1 - Y0


def process_image(input_filename, border_pixels):
    processed = Image.new(mode="RGB", size=(WIDTH_PX, HEIGHT_PX), color=(255, 255, 255))
    try:
        with Image.open(input_filename, "r") as im:
            width, height = im.width, im.height
            if im.height > im.width:
                im = im.rotate(angle=270, expand=True)
                width, height = height, width
            scale_x = (W - 2 * border_pixels) / width
            scale_y = (H - 2 * border_pixels) / height
            scale = min(scale_x, scale_y)
            w = int(scale * width)
            h = int(scale * height)
            offset_x = int((WIDTH_PX - w) / 2)
            offset_y = int((HEIGHT_PX - h) / 2)
            im = im.resize(size=(w, h), resample=Image.Resampling.NEAREST)
            processed.paste(im, box=(offset_x, offset_y))
            return processed
    except UnidentifiedImageError:
        print(f"Unsupported image file \"{input_filename}\"")

def output_image(im, output_filename, print_document_name="image"):
    if len(output_filename) > 0:
        im.save(output_filename, dpi=(DPI, DPI))
    else:
        printer_name = get_printer_name()
        print_image(printer_name, print_document_name, im)


def main():
    parser = argparse.ArgumentParser(
        description="Adjust image to print on Canon Selphy without cropping, print it or save to file")
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
                        help="(optional) Output filename or directory for processed images."
                             "If not provided the image will be sent to the printer",
                        type=str,
                        default="",
                        required=False)
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Input path \"{args.input}\" does not exist.")
        exit(1)

    if len(args.output) > 0:
        output_parent = os.path.dirname(args.output)
        if not os.path.isdir(output_parent):
            print(f"Parent directory \"{output_parent}\" must exist.")
            exit(2)

    if args.border > MAX_BORDER:
        print(f"Border cannot exceed {MAX_BORDER} mm.")
        exit(3)

    border_px = int(args.border / INCH * DPI)

    if os.path.isfile(args.input):
        result = process_image(args.input, border_px)
        document_name = f"Image \"{args.input}\""
        output_image(result, args.output, document_name)
    else:
        root, subdirectory, files = next(os.walk(args.input))
        for filename in files:
            fn = os.path.join(root, filename)
            basename, ext = os.path.splitext(filename)
            fn_output = os.path.join(args.output, basename + "-print" + ext)
            result = process_image(fn, border_px)
            output_image(result, fn_output, filename)


if __name__ == "__main__":
    main()
