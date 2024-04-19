# selphyprint

`selphyprints` prepars image files for printing on Canon Selphy printers without cropping. 
The image is resized to fit into the usable area of the print and padded with borders required to avoid cropping.
Additional borders can be added if desired. The resulting image is saved into a file that can be sent directly to the printer.

# Printer configuration

Configure your printer to use the borderless mode and from your system print dialog choose "Fill Page".

# Installation

Install from PiPY:
```commandline
pip install selphyprint
```

# Usage

To prepare a single image for printing use the following command:
```commandline
selphyprint --input <input-image-file> --output <output-image-file>
```
For example:
```commandline
selphyprint --input waves.tif --output waves-print.tif
```

You can also process all files in a directory, for example:
```commandline
selphyprint --input /home/roman/photos --output /home/roman/prints
```
